using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Wsh.XLS.Editor {

    public class XLSBuilder {

        private static List<List<object>> GetValues(ExcelWorksheet worksheet, int totalColumnNumber) {
            List<List<object>> list = new List<List<object>>();//按列存储数据
            // 读取数据行
            // Log.Info("Row", worksheet.Dimension.End.Row, "Column", worksheet.Dimension.End.Column, totalColumnNumber, "Cells.Rows", worksheet.Cells.Rows, "Cells.Columns", worksheet.Cells.Columns);
            for(int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++) {
                for(int colNum = 1; colNum <= worksheet.Dimension.End.Column; colNum++) {
                    var data = worksheet.Cells[rowNum, colNum];
                    object value = XLSDefine.NULL_CHAR;
                    if(data.Value != null) {
                        value = data.Value;
                    }
                    Log.Info(rowNum, colNum, value);
                }

            }

            return list;
        }

        private static Dictionary<int, XLSHeadInfo> GetHeadInfoDic(ExcelWorksheet worksheet, ref int idColIndex, ref int valueColIndex) {
            //以列为key
            Dictionary<int, XLSHeadInfo> headInfoDic = new Dictionary<int, XLSHeadInfo>();
            //以名称为key
            Dictionary<string, int> nameDic = new Dictionary<string, int>();
            // 读取列名，
            for(int i = 1; i <= worksheet.Dimension.End.Column; i++) {
                string headText = worksheet.Cells[1, i].Value.ToString();
                string[] headData = headText.Split(':');
                if(headData[0] == XLSDefine.ID_CHAR) {
                    idColIndex = i;
                    nameDic.Add(headData[0], 1);
                    headInfoDic.Add(i, new XLSHeadInfo(i, headData[0], XLSFieldType.STRING));
                } else if(headData[0] == XLSDefine.VALUE_CHAR) {
                    valueColIndex = i;
                    nameDic.Add(headData[0], 1);
                    headInfoDic.Add(i, new XLSHeadInfo(i, headData[0], XLSFieldType.INT));
                } else {
                    if(headData.Length < 2) {
                        throw new Exception($"' {headData[0]} ' no value-type in (1, {i}).");
                    }
                    if(!XLSFieldType.Contain(headData[1])) {
                        throw new Exception($"' {headText} ' no define value-type in (1, {i}). {XLSFieldType.GetAllDefineType()}");
                    }
                    if(nameDic.ContainsKey(headData[0])) {
                        throw new Exception($"' {headText} ' exist same name in (1, {i}).");
                    }
                    nameDic.Add(headData[0], 1);
                    headInfoDic.Add(i, new XLSHeadInfo(i, headData[0], headData[1]));
                }
            }
            if(!nameDic.ContainsKey(XLSDefine.ID_CHAR)) {
                throw new Exception($"Do not define ' {XLSDefine.ID_CHAR} '.");
            }
            if(!nameDic.ContainsKey(XLSDefine.VALUE_CHAR)) {
                throw new Exception($"Do not define ' {XLSDefine.VALUE_CHAR} '.");
            }
            return headInfoDic;
        }

        private static Dictionary<int, XLSIDInfo> GetIDInfoDic(ExcelWorksheet worksheet, int idColIndex, int valueColIndex) {
            //以 VALUE(int) 为key
            Dictionary<int, XLSIDInfo> idInfoDic = new Dictionary<int, XLSIDInfo>();
            //以 ID(string) 为key
            Dictionary<string, int> idDic = new Dictionary<string, int>();
            int idHeight = 0;
            int lastVauleId = 0;
            for(int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++) {
                var idCellData = worksheet.Cells[rowNum, idColIndex];
                var valueCellData = worksheet.Cells[rowNum, valueColIndex];
                if(idCellData.Value != null) {
                    if(valueCellData.Value == null) {
                        throw new Exception($"' {idCellData.Value} ' do not define ' {XLSDefine.VALUE_CHAR} '.");
                    }
                    string id = idCellData.Value.ToString();
                    int value = 0;
                    int.TryParse(valueCellData.Value.ToString(), out value);
                    if(value == 0) {
                        throw new Exception($"' {idCellData.Value} '  -- ' {XLSDefine.VALUE_CHAR} ' must be ' int '.");
                    }
                    if(lastVauleId != 0) {
                        idInfoDic[lastVauleId].SetHeight(idHeight);
                    }
                    idHeight = 1;
                    if(idInfoDic.ContainsKey(value)) {
                        throw new Exception($"' {idCellData.Value} '  -- ' {XLSDefine.VALUE_CHAR} ' exist same value.");
                    }
                    if(idDic.ContainsKey(id)) {
                        throw new Exception($"' {idCellData.Value} ' exist same id.");
                    }
                    idDic.Add(id, 1);
                    idInfoDic.Add(value, new XLSIDInfo(id, value, rowNum));
                    lastVauleId = value;
                } else {
                    idHeight++;
                    if(rowNum == worksheet.Dimension.End.Row) {
                        idInfoDic[lastVauleId].SetHeight(idHeight);
                    }
                }
            }
            return idInfoDic;
        }

        public static void BuildFile(string fileName, string filePath, string outputFilePath) {
            using(FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                using(var package = new ExcelPackage(fs)) {
                    // 获取第一个工作表
                    var worksheet = package.Workbook.Worksheets[0];
                    Dictionary<int, XLSHeadInfo> headInfoDic;
                    Dictionary<int, XLSIDInfo> idInfoDic;
                    int idColIndex = 0;
                    int valueColIndex = 0;
                    try {
                        headInfoDic = GetHeadInfoDic(worksheet, ref idColIndex, ref valueColIndex);
                    } catch(Exception e) {
                        throw new Exception(e.Message + " from " + filePath);
                    }

                    try {
                        idInfoDic = GetIDInfoDic(worksheet, idColIndex, valueColIndex);
                        //foreach(var data in idInfoDic.Values) {
                        //    Log.Info("id", data.Id, "value", data.Value, "rowIndex", data.RowIndex, "height", data.Height);
                        //}
                    } catch(Exception e) {
                        throw new Exception(e.Message + " from " + filePath);
                    }

                    try {
                        XLSClassGenerator.Generate(headInfoDic, idInfoDic, worksheet, fileName, outputFilePath);
                    } catch(Exception e) {
                        throw new Exception(e.Message + " from " + filePath);
                    }

                }
            }
        }


    }
}

