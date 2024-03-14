using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using UnityEditor;

namespace Wsh.Xlsx.Editor {

    public class XlsxBuilder {

        private static List<List<object>> GetValues(ExcelWorksheet worksheet, int totalColumnNumber) {
            List<List<object>> list = new List<List<object>>();//按列存储数据
            // 读取数据行
            // Log.Info("Row", worksheet.Dimension.End.Row, "Column", worksheet.Dimension.End.Column, totalColumnNumber, "Cells.Rows", worksheet.Cells.Rows, "Cells.Columns", worksheet.Cells.Columns);
            for(int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++) {
                for(int colNum = 1; colNum <= worksheet.Dimension.End.Column; colNum++) {
                    var data = worksheet.Cells[rowNum, colNum];
                    object value = XlsxDefine.NULL_CHAR;
                    if(data.Value != null) {
                        value = data.Value;
                    }
                    Log.Info(rowNum, colNum, value);
                }

            }

            return list;
        }

        private static Dictionary<int, XlsxHeadInfo> GetHeadInfoDic(ExcelWorksheet worksheet, ref int idColIndex, ref int valueColIndex) {
            //以列为key
            Dictionary<int, XlsxHeadInfo> headInfoDic = new Dictionary<int, XlsxHeadInfo>();
            //以名称为key
            Dictionary<string, int> nameDic = new Dictionary<string, int>();
            // 读取列名，
            for(int i = 1; i <= worksheet.Dimension.End.Column; i++) {
                string headText = worksheet.Cells[1, i].Value.ToString();
                string[] headData = headText.Split(':');
                if(headData[0] == XlsxDefine.ID_CHAR) {
                    idColIndex = i;
                    nameDic.Add(headData[0], 1);
                    headInfoDic.Add(i, new XlsxHeadInfo(i, headData[0], XlsxFieldType.STRING));
                } else if(headData[0] == XlsxDefine.VALUE_CHAR) {
                    valueColIndex = i;
                    nameDic.Add(headData[0], 1);
                    headInfoDic.Add(i, new XlsxHeadInfo(i, headData[0], XlsxFieldType.INT));
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
                    headInfoDic.Add(i, new XlsxHeadInfo(i, headData[0], headData[1]));
                }
            }
            if(!nameDic.ContainsKey(XlsxDefine.ID_CHAR)) {
                throw new Exception($"Do not define ' {XlsxDefine.ID_CHAR} '.");
            }
            if(!nameDic.ContainsKey(XlsxDefine.VALUE_CHAR)) {
                throw new Exception($"Do not define ' {XlsxDefine.VALUE_CHAR} '.");
            }
            return headInfoDic;
        }

        private static Dictionary<int, XlsxIDInfo> GetIDInfoDic(ExcelWorksheet worksheet, int idColIndex, int valueColIndex) {
            //以 VALUE(int) 为key
            Dictionary<int, XlsxIDInfo> idInfoDic = new Dictionary<int, XlsxIDInfo>();
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
                        idInfoDic[lastVauleId].SetHeight(idHeight);
                    }
                    idHeight = 1;
                    if(idInfoDic.ContainsKey(value)) {
                        throw new Exception($"' {idCellData.Value} '  -- ' {XlsxDefine.VALUE_CHAR} ' exist same value.");
                    }
                    if(idDic.ContainsKey(id)) {
                        throw new Exception($"' {idCellData.Value} ' exist same id.");
                    }
                    idDic.Add(id, 1);
                    idInfoDic.Add(value, new XlsxIDInfo(id, value, rowNum));
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
                    Dictionary<int, XlsxHeadInfo> headInfoDic;
                    Dictionary<int, XlsxIDInfo> idInfoDic;
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
                        XlsxClassGenerator.Generate(headInfoDic, idInfoDic, worksheet, fileName, outputFilePath);
                    } catch(Exception e) {
                        throw new Exception(e.Message + " from " + filePath);
                    }

                }
            }
        }

        public static void BuildFolder(string xlsxFolder, string outputFolder) {
            DirectoryInfo dir = new DirectoryInfo(xlsxFolder);
            FileInfo[] fileInfos = dir.GetFiles();
            EditorUtility.DisplayProgressBar("Generate xlsx Csharp file", "start generate...", 0);
            int currentIndex = 0;
            float totalNumber = fileInfos.Length;
            try {
                foreach(FileInfo fileInfo in fileInfos) {
                    string filePath = fileInfo.FullName;
                    string fileName = Path.GetFileNameWithoutExtension(fileInfo.Name);
                    string extension = fileInfo.Extension;
                    //Log.Info(filePath, fileName, extension);
                    currentIndex++;
                    EditorUtility.DisplayProgressBar("Generate xlsx Csharp file", "generate " + fileInfo.Name, currentIndex / totalNumber);
                    if(extension != XlsxDefine.XLSX_EXTENSION) {
                        continue;
                    }
                    if(extension == XlsxDefine.XLS_EXTENSION) {
                        Log.Error("Dont support '.xls'. Please use '.xlsx'");
                        continue;
                    }
                    BuildFile(fileName, filePath, Path.Combine(outputFolder, fileName + XlsxDefine.CLASS_SUFFIX + ".cs"));
                }
                EditorUtility.ClearProgressBar();
                EditorUtility.DisplayDialog("Generate xlsx", "Generate Success", "Ok");
            } catch(Exception e) {
                EditorUtility.ClearProgressBar();
                EditorUtility.DisplayDialog("Generate xlsx", "Generate Failed", "Ok");
                throw e;
            }

            AssetDatabase.Refresh();
        }

        public static void ClearFolder(string folder) {
            // 检查路径是否存在并且是个文件夹
            if(!Directory.Exists(folder)) {
                Console.WriteLine($"'{folder}' not exist.");
                return;
            }

            // 获取该文件夹下的所有文件和子文件夹
            foreach(var filePath in Directory.GetFiles(folder)) {
                File.Delete(filePath);
            }
            // 清空子文件夹中的文件（递归调用）
            foreach(var subFolderPath in Directory.GetDirectories(folder)) {
                ClearFolder(subFolderPath);
                // 如果不需要保留子文件夹结构，还可以删除子文件夹
                // Directory.Delete(subFolderPath, true); // 第二个参数为true表示递归删除
            }
            //Directory.Delete(folder, true);
            AssetDatabase.Refresh();
        }


    }
}

