using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace Wsh.Xlsx.Editor {

    public class XlsxGenerateData : IDisposable {

        public Dictionary<int, XlsxHeadInfo> HeadInfoDic => m_headInfoDic;
        public Dictionary<int, XlsxIDInfo> IdInfoDic => m_idInfoDic;
        public object[,] Content => m_content;
        public string XlsxName => m_xlsxName;
        public string OutputFilePath => m_outputFilePath;

        //以列为key
        private Dictionary<int, XlsxHeadInfo> m_headInfoDic;
        //以 VALUE(int) 为key
        private Dictionary<int, XlsxIDInfo> m_idInfoDic;
        private string m_xlsxName;
        private string m_outputFilePath;
        private object[,] m_content;

        public XlsxGenerateData(string xlsxName, string outputFilePath) {
            m_headInfoDic = new Dictionary<int, XlsxHeadInfo>();
            m_idInfoDic = new Dictionary<int, XlsxIDInfo>();
            m_xlsxName = xlsxName;
            m_outputFilePath = outputFilePath;
        }

        public void Init(ExcelWorksheet worksheet) {
            int idColIndex = 0;
            int valueColIndex = 0;
            try {
                SetHeadInfoDic(worksheet, ref idColIndex, ref valueColIndex);
            } catch(Exception e) {
                throw new Exception(e.Message + " from " + XlsxName);
            }
            try {
                SetIDInfoDic(worksheet, idColIndex, valueColIndex);
                //foreach(var data in idInfoDic.Values) {
                //    Log.Info("id", data.Id, "value", data.Value, "rowIndex", data.RowIndex, "height", data.Height);
                //}
            } catch(Exception e) {
                throw new Exception(e.Message + " from " + XlsxName);
            }
            SetValues(worksheet);
        }

        private void SetHeadInfoDic(ExcelWorksheet worksheet, ref int idColIndex, ref int valueColIndex) {
            //以名称为key
            Dictionary<string, int> nameDic = new Dictionary<string, int>();
            // 读取列名，
            for(int i = 1; i <= worksheet.Dimension.End.Column; i++) {
                string headText = worksheet.Cells[1, i].Value.ToString();
                string[] headData = headText.Split(':');
                if(headData[0] == XlsxDefine.ID_CHAR) {
                    idColIndex = i;
                    nameDic.Add(headData[0], 1);
                    m_headInfoDic.Add(i, new XlsxHeadInfo(i, headData[0], XlsxFieldType.STRING));
                } else if(headData[0] == XlsxDefine.VALUE_CHAR) {
                    valueColIndex = i;
                    nameDic.Add(headData[0], 1);
                    m_headInfoDic.Add(i, new XlsxHeadInfo(i, headData[0], XlsxFieldType.INT));
                } else {
                    if(headData.Length < 2) {
                        throw new Exception($"' {headData[0]} ' no value-type in (1, {i}).");
                    }
                    if(!XlsxFieldType.Contain(headData[1])) {
                        throw new Exception($"' {headText} ' no define value-type in (1, {i}). {XlsxFieldType.GetAllDefineType()}");
                    }
                    if(nameDic.ContainsKey(headData[0])) {
                        throw new Exception($"' {headText} ' exist same name in (1, {i}).");
                    }
                    nameDic.Add(headData[0], 1);
                    m_headInfoDic.Add(i, new XlsxHeadInfo(i, headData[0], headData[1]));
                }
            }
            if(!nameDic.ContainsKey(XlsxDefine.ID_CHAR)) {
                throw new Exception($"Do not define ' {XlsxDefine.ID_CHAR} '.");
            }
            if(!nameDic.ContainsKey(XlsxDefine.VALUE_CHAR)) {
                throw new Exception($"Do not define ' {XlsxDefine.VALUE_CHAR} '.");
            }
        }

        private void SetIDInfoDic(ExcelWorksheet worksheet, int idColIndex, int valueColIndex) {
            //以 ID(string) 为key
            Dictionary<string, int> idDic = new Dictionary<string, int>();
            int idHeight = 0;
            int lastVauleId = 0;
            for(int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++) {
                var idCellData = worksheet.Cells[rowNum, idColIndex];
                var valueCellData = worksheet.Cells[rowNum, valueColIndex];
                if(idCellData.Value != null) {
                    if(valueCellData.Value == null) {
                        throw new Exception($"' {idCellData.Value} ' do not define ' {XlsxDefine.VALUE_CHAR} '.");
                    }
                    string id = idCellData.Value.ToString();
                    int value = 0;
                    int.TryParse(valueCellData.Value.ToString(), out value);
                    if(value == 0) {
                        throw new Exception($"' {idCellData.Value} '  -- ' {XlsxDefine.VALUE_CHAR} ' must be ' int '.");
                    }
                    if(lastVauleId != 0) {
                        m_idInfoDic[lastVauleId].SetHeight(idHeight);
                    }
                    idHeight = 1;
                    if(m_idInfoDic.ContainsKey(value)) {
                        throw new Exception($"' {idCellData.Value} '  -- ' {XlsxDefine.VALUE_CHAR} ' exist same value.");
                    }
                    if(idDic.ContainsKey(id)) {
                        throw new Exception($"' {idCellData.Value} ' exist same id.");
                    }
                    idDic.Add(id, 1);
                    m_idInfoDic.Add(value, new XlsxIDInfo(id, value, rowNum));
                    lastVauleId = value;
                } else {
                    idHeight++;
                    if(rowNum == worksheet.Dimension.End.Row) {
                        m_idInfoDic[lastVauleId].SetHeight(idHeight);
                    }
                }
            }
        }

        private void SetValues(ExcelWorksheet worksheet) {
            // 读取数据行
            // Log.Info("Row", worksheet.Dimension.End.Row, "Column", worksheet.Dimension.End.Column, totalColumnNumber, "Cells.Rows", worksheet.Cells.Rows, "Cells.Columns", worksheet.Cells.Columns);
            m_content = new object[worksheet.Dimension.End.Row + 1, worksheet.Dimension.End.Column + 1];
            for(int rowNum = 1; rowNum <= worksheet.Dimension.End.Row; rowNum++) {
                for(int colNum = 1; colNum <= worksheet.Dimension.End.Column; colNum++) {
                    var data = worksheet.Cells[rowNum, colNum];
                    m_content[rowNum, colNum] = data.Value;
                    /*object value = XlsxDefine.NULL_CHAR;
                    if(data.Value != null) {
                        value = data.Value;
                    }
                    Log.Info(rowNum, colNum, value);*/
                }
            }
        }

        public void Dispose() {
            foreach(var headInfo in m_headInfoDic.Values) {
                headInfo.Dispose();
            }
            m_headInfoDic.Clear();
            m_headInfoDic = null;
            foreach(var idInfo in m_idInfoDic.Values) {
                idInfo.Dispose();
            }
            m_idInfoDic.Clear();
            m_idInfoDic = null;
            m_content = null;
        }
    }
}