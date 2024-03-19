using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Codice.Client.BaseCommands.Fileinfo;
using OfficeOpenXml;
using UnityEditor;

namespace Wsh.Xlsx.Editor {

    public class XlsxBuilder {

        public static void BuildFile(string fileName, string filePath, string outputFolder, ref Dictionary<string, int> idDic) {
            using(FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                using(var package = new ExcelPackage(fs)) {
                    // 获取第一个工作表
                    var worksheet = package.Workbook.Worksheets[0];
                    string outputFilePath = Path.Combine(outputFolder, fileName + XlsxDefine.CLASS_SUFFIX + ".cs");
                    using(XlsxGenerateData xlsxGenerateData = new XlsxGenerateData(fileName, outputFilePath)) {
                        try {
                            xlsxGenerateData.Init(worksheet);
                            if(fileName != XlsxDefine.LOCAL_FILE_NAME) {
                                var idDicTemp = xlsxGenerateData.GetIds();
                                foreach(var key in idDicTemp.Keys) {
                                    if(idDic.ContainsKey(key)) {
                                        throw new Exception($"Exist same key in '{filePath}' with other xlsx.");
                                    } else {
                                        idDic.Add(key, idDicTemp[key]);
                                    }
                                }
                            } else {
                                XlsxClassGenerator.GenerateLocalization(xlsxGenerateData.GetIds(), outputFolder);
                            }
                        } catch(Exception e) {
                            throw e;
                        }
                        try {
                            XlsxClassGenerator.Generate(xlsxGenerateData);
                        } catch(Exception e) {
                            throw new Exception(e.Message + " from " + filePath);
                        }
                    }
                }
            }
        }

        public static void BuildFolder(string xlsxFolder, string outputFolder) {
            DirectoryInfo dir = new DirectoryInfo(xlsxFolder);
            System.IO.FileInfo[] fileInfos = dir.GetFiles();
            EditorUtility.DisplayProgressBar("Generate xlsx Csharp file", "start generate...", 0);
            int currentIndex = 0;
            float totalNumber = fileInfos.Length;
            Dictionary<string, int> idDic = new Dictionary<string, int>();
            idDic.Add(XlsxDefine.NONE, 0);
            try {
                foreach(System.IO.FileInfo fileInfo in fileInfos) {
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
                    BuildFile(fileName, filePath, outputFolder, ref idDic);
                }
                EditorUtility.DisplayProgressBar("Generate xlsx Csharp file", "generate " + XlsxDefine.XLSX_ID_FILE_NAME, 0);
                XlsxClassGenerator.GenerateXlsxIdListFile(idDic, outputFolder);
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

