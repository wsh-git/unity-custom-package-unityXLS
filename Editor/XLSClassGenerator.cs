using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;

namespace Wsh.XLS.Editor {

    public static class XLSClassGenerator {

        private static void AddTab(StringBuilder stringBuilder, ref int tabIndex) {
            tabIndex++;
            if(XLSDefine.TAB_CHAR_ARRAY.Length <= tabIndex ) {
                Log.Error("XLSDefine.TAB_CHAR_ARRAY outside the bounds of array.", tabIndex);
            }
            stringBuilder.Append(XLSDefine.TAB_CHAR_ARRAY[tabIndex]);
        }

        private static void SubTab(StringBuilder stringBuilder, ref int tabIndex) {
            tabIndex--;
            if(tabIndex <= 0) {
                Log.Error("XLSDefine.TAB_CHAR_ARRAY index: 0.");
            }
            stringBuilder.Append(XLSDefine.TAB_CHAR_ARRAY[tabIndex]);
        }

        private static void KeepTap(StringBuilder stringBuilder, ref int tabIndex) {
            stringBuilder.Append(XLSDefine.TAB_CHAR_ARRAY[tabIndex]);
        }

        public static void Generate(Dictionary<int, XLSHeadInfo> headInfoDic, Dictionary<int, XLSIDInfo> idInfoDic, ExcelWorksheet worksheet, string xlsName, string outputFilePath) {
            try {
                StringBuilder stringBuilder = new StringBuilder();
                stringBuilder.Append("// Automatically generated. Do not modify it manually!!!\n\n");
                stringBuilder.AppendLine("using System.Collections.Generic;");
                stringBuilder.AppendLine("using System.Linq;\n");
                stringBuilder.Append("namespace Wsh.XLS {\n\n");
                int tabIndex = 0;
                AddTab(stringBuilder, ref tabIndex);
                stringBuilder.Append($"public static class {xlsName}{XLSDefine.CLASS_SUFFIX} " + "{\n\n");
                string structName = xlsName + XLSDefine.STRUCT_SUFFIX;
                stringBuilder.Append(CreateStruct(xlsName, structName, headInfoDic, ref tabIndex));
                stringBuilder.Append(CreateDictionary(headInfoDic, idInfoDic, worksheet, structName, ref tabIndex));
                stringBuilder.Append(CreateFunctions(headInfoDic, structName, ref tabIndex));
                stringBuilder.Append(XLSDefine.TAB_CHAR_ARRAY[1] + "}\n\n");
                stringBuilder.Append("}\n");
                File.WriteAllText(outputFilePath, stringBuilder.ToString());
                Log.Info($"{xlsName}{XLSDefine.CLASS_SUFFIX}.cs generate success.");
            } catch(Exception e) {
                throw e;
            }
        }

        private static string CreateStruct(string xlsName, string structName, Dictionary<int, XLSHeadInfo> headInfoDic, ref int tabIndex) {
            StringBuilder stringBuilder = new StringBuilder();
            AddTab(stringBuilder, ref tabIndex);
            stringBuilder.Append($"public struct {structName} " + "{\n");
            tabIndex++;
            foreach(var value in headInfoDic.Values) {
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.Append($"public {XLSFieldType.GetValueTypeString(value.ValueType)} {value.Name};\n");
            }
            SubTab(stringBuilder, ref tabIndex);
            stringBuilder.Append("}\n\n");
            return stringBuilder.ToString();
        }

        private static string CreateDictionary(Dictionary<int, XLSHeadInfo> headInfoDic, Dictionary<int, XLSIDInfo> idInfoDic, ExcelWorksheet worksheet, string structName, ref int tabIndex) {
            try {
                StringBuilder stringBuilder = new StringBuilder();
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.Append($"private readonly static Dictionary<int, {structName}> m_dic = new Dictionary<int, {structName}> " + "{\n");
                tabIndex++;
                foreach(var value in idInfoDic.Values) {
                    KeepTap(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("{" + $" {value.Value}, new {structName} " + "{");
                    tabIndex++;
                    foreach(var headInfo in headInfoDic.Values) {
                        KeepTap(stringBuilder, ref tabIndex);
                        stringBuilder.AppendLine($"{headInfo.Name} = {GetValue(value, headInfo, worksheet, ref tabIndex)},");
                    }
                    SubTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("}},");
                }
                SubTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine("};\n");
                return stringBuilder.ToString();
            } catch(Exception e) {
                throw e;
            }
        }

        private static string GetValue(XLSIDInfo idInfo, XLSHeadInfo headInfo, ExcelWorksheet worksheet, ref int tabIndex) {
            try {
                object obj = worksheet.Cells[idInfo.RowIndex, headInfo.Index].Value;
                switch(headInfo.ValueType) {
                    case XLSFieldType.STRING:
                        return TryConvertString(obj);
                    case XLSFieldType.STRING_ARRAY:
                        return TryConvertStringArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.STRING_2D_ARRAY:
                        return TryConvertString2DArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);;
                    case XLSFieldType.INT:
                        return TryConvertInt(obj);
                    case XLSFieldType.INT_ARRAY:
                        return TryConvertIntArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.INT_2D_ARRAY:
                        return TryConvertInt2DArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.INT_TEN:
                        return TryConvertIntTen(obj);
                    case XLSFieldType.INT_TEN_ARRAY:
                        return TryConvertIntTenArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.INT_TEN_2D_ARRAY:
                        return TryConvertIntTen2DArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.INT_HUNDRED:
                        return TryConvertIntHundred(obj);
                    case XLSFieldType.INT_HUNDRED_ARRAY:
                        return TryConvertIntHundredArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.INT_HUNDRED_2D_ARRAY:
                        return TryConvertIntHundred2DArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.INT_THOUSAND:
                        return TryConvertIntThousand(obj);
                    case XLSFieldType.INT_THOUSAND_ARRAY:
                        return TryConvertIntThousandArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.INT_THOUSAND_2D_ARRAY:
                        return TryConvertIntThousand2DArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.BOOLEAN:
                        return TryConvertBoolean(obj);
                    case XLSFieldType.BOOLEAN_ARRAY:
                        return TryConvertBooleanArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                    case XLSFieldType.BOOLEAN_2D_ARRAY:
                        return TryConvertBoolean2DArray(idInfo.RowIndex, headInfo.Index, worksheet, idInfo.Height, ref tabIndex);
                }
                return "null";
            } catch(Exception e) {
                throw new Exception(e.Message + $" in ({idInfo.RowIndex}, {headInfo.Index}). ");
            }
        }
        private static string TryConvertIntTenArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.FLOAT_ARRAY, ref tabIndex, TryConvertIntTen);
        }
        private static string TryConvertIntHundredArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.FLOAT_ARRAY, ref tabIndex, TryConvertIntHundred);
        }

        private static string TryConvertIntThousandArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.FLOAT_ARRAY, ref tabIndex, TryConvertIntThousand);
        }

        private static string TryConvertStringArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.STRING_ARRAY, ref tabIndex, TryConvertString);
        }

        private static string TryConvertIntArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.INT_ARRAY, ref tabIndex, TryConvertInt);
        }

        private static string TryConvertBooleanArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.BOOLEAN_ARRAY, ref tabIndex, TryConvertBoolean);
        }

        private static string TryConvertArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, string valueType, ref int tabIndex, Func<object, string> tryConvertHandler) {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"new {valueType} " + "{");
            bool hasContent = false;
            tabIndex++;
            for(int row = rowIndex; row < rowIndex+maxHeight; row++) {
                if(worksheet.Cells[row, colIndex].Value == null) {
                    break;
                }
                hasContent = true;
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.Append(tryConvertHandler(worksheet.Cells[row, colIndex].Value));
                stringBuilder.AppendLine(",");
            }
            SubTab(stringBuilder, ref tabIndex);
            stringBuilder.Append("}");
            if(hasContent) {
                return stringBuilder.ToString();
            }
            return "null";
        }

        private static string TryConvertString2DArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.STRING, ref tabIndex, TryConvertString);
        }
        
        private static string TryConvertInt2DArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.INT, ref tabIndex, TryConvertInt);
        }
        private static string TryConvertIntTen2DArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.FLOAT, ref tabIndex, TryConvertIntTen);
        }

        private static string TryConvertIntHundred2DArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.FLOAT, ref tabIndex, TryConvertIntHundred);
        }

        private static string TryConvertIntThousand2DArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.FLOAT, ref tabIndex, TryConvertIntThousand);
        }

        private static string TryConvertBoolean2DArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, worksheet, maxHeight, XLSFieldType.BOOLEAN, ref tabIndex, TryConvertBoolean);
        }
        
        private static string TryConvert2DArray(int rowIndex, int colIndex, ExcelWorksheet worksheet, int maxHeight, string valueType, ref int tabIndex, Func<object, string> tryConvertHandler) {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"new {valueType}[][] " + "{");
            bool hasContent = false;
            tabIndex++;
            for(int row = rowIndex; row < rowIndex + maxHeight; row++) {
                if(worksheet.Cells[row, colIndex].Value == null) {
                    break;
                }
                hasContent = true;
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"new {valueType}[] " + "{");
                string text = worksheet.Cells[row, colIndex].Value.ToString();
                string[] arrayText = text.Split('#');
                tabIndex++;
                for(int i = 0; i < arrayText.Length; i++) {
                    KeepTap(stringBuilder, ref tabIndex);
                    stringBuilder.Append(tryConvertHandler(arrayText[i]));
                    stringBuilder.AppendLine(",");
                }
                SubTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine("},");
            }
            SubTab(stringBuilder, ref tabIndex);
            stringBuilder.Append("}");
            if(hasContent) {
                return stringBuilder.ToString();
            } else {
                return "null";
            }
        }

        private static string TryConvertIntTen(object obj) {
            return TryConvertFloat(obj, 10f, "F1");
        }

        private static string TryConvertIntHundred(object obj) {
            return TryConvertFloat(obj, 100f, "F2");
        }

        private static string TryConvertIntThousand(object obj) {
            return TryConvertFloat(obj, 1000f, "F3");
        }
        
        private static string TryConvertFloat(object obj, float ratio, string numberFormat) {
            if(obj == null) {
                return "0f";
            } else {
                int value = 0;
                try {
                    value = int.Parse(obj.ToString());
                    if(numberFormat == "F3") {
                        return $"{value/ratio:F3}f";
                    } else if(numberFormat == "F2") {
                         return $"{value/ratio:F2}f";
                    } else {
                        return $"{value/ratio:F1}f";
                    }
                } catch(Exception e) {
                    throw e;
                }
            }
        }

        private static string TryConvertInt(object obj) {
            int result = 0;
            if(obj != null) {
                try {
                    result = int.Parse(obj.ToString());
                } catch(Exception e) {
                    throw e;
                }
            }
            return result.ToString();
        }
        
        private static string TryConvertBoolean(object obj) {
            bool result = false;
            if(obj != null) {
               try {
                    result = bool.Parse(obj.ToString());
                } catch(Exception e) {
                    throw e;
                }
            }
            return result.ToString().ToLower();
        }

        private static string TryConvertString(object obj) {
            if(obj == null) {
                return "\"\"";
            } else {
                return "\"" + obj.ToString() + "\"";
            }
        }

        private static string CreateFunctions(Dictionary<int, XLSHeadInfo> headInfoDic, string structName, ref int tabIndex) {
            StringBuilder stringBuilder = new StringBuilder();

            KeepTap(stringBuilder, ref tabIndex);
            stringBuilder.AppendLine($"public static List<{structName}> GetDataList() " + "{");
            AddTab(stringBuilder, ref tabIndex);
            stringBuilder.AppendLine("return m_dic.Values.ToList();");
            SubTab(stringBuilder, ref tabIndex);
            stringBuilder.AppendLine("}\n");

            KeepTap(stringBuilder, ref tabIndex);
            stringBuilder.AppendLine("public static bool Contain(int id) {");
            AddTab(stringBuilder, ref tabIndex);
            stringBuilder.AppendLine("return m_dic.ContainsKey(id);");
            SubTab(stringBuilder,ref tabIndex);
            stringBuilder.AppendLine("}\n");

            foreach(var headInfo in headInfoDic.Values) {
                if(headInfo.Name == XLSDefine.ID_CHAR || headInfo.Name == XLSDefine.VALUE_CHAR) {
                    continue;
                }

                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"public static {XLSFieldType.GetValueTypeString(headInfo.ValueType)} {headInfo.Name}(int id) " + "{");
                AddTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine("if(Contain(id)) {");
                AddTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"return m_dic[id].{headInfo.Name};");
                SubTab(stringBuilder,ref tabIndex);
                stringBuilder.AppendLine("}");
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"return {XLSFieldType.GetDefaultValue(headInfo.ValueType)};");
                SubTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine("}\n");

                string returnType = "";
                bool isArray = XLSFieldType.IsArray(headInfo.ValueType, ref returnType);
                if(isArray) {
                    KeepTap(stringBuilder,ref tabIndex);
                    stringBuilder.AppendLine($"public static {returnType} {headInfo.Name}FromArray(int id, int index)" + "{");
                    AddTab(stringBuilder,ref tabIndex);
                    stringBuilder.AppendLine("if(Contain(id)) {");
                    AddTab(stringBuilder,ref tabIndex);
                    stringBuilder.AppendLine($"{XLSFieldType.GetValueTypeString(headInfo.ValueType)} array = {headInfo.Name}(id);");
                    KeepTap(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("if(array.Length > index) {");
                    AddTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("return array[index];");
                    SubTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("}");
                    SubTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("}");
                    KeepTap(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine($"return {XLSFieldType.GetDefaultValue(returnType)};");
                    SubTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("}\n");
                }

            }
            return stringBuilder.ToString();
        }

    }
}