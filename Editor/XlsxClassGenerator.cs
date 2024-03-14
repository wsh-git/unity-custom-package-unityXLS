using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Wsh.Xlsx.Editor {

    public static class XlsxClassGenerator {

        private static void AddTab(StringBuilder stringBuilder, ref int tabIndex) {
            tabIndex++;
            if(XlsxDefine.TAB_CHAR_ARRAY.Length <= tabIndex ) {
                Log.Error("XlsxDefine.TAB_CHAR_ARRAY outside the bounds of array.", tabIndex);
            }
            stringBuilder.Append(XlsxDefine.TAB_CHAR_ARRAY[tabIndex]);
        }

        private static void SubTab(StringBuilder stringBuilder, ref int tabIndex) {
            tabIndex--;
            if(tabIndex <= 0) {
                Log.Error("XlsxDefine.TAB_CHAR_ARRAY index: 0.");
            }
            stringBuilder.Append(XlsxDefine.TAB_CHAR_ARRAY[tabIndex]);
        }

        private static void KeepTap(StringBuilder stringBuilder, ref int tabIndex) {
            stringBuilder.Append(XlsxDefine.TAB_CHAR_ARRAY[tabIndex]);
        }

        public static void Generate(XlsxGenerateData xlsxGenerateData) {
            try {
                StringBuilder stringBuilder = new StringBuilder();
                stringBuilder.Append("// Automatically generated. Do not modify it manually!!!\n\n");
                stringBuilder.AppendLine("using System.Collections.Generic;");
                stringBuilder.AppendLine("using System.Linq;\n");
                stringBuilder.Append("namespace Wsh.Xlsx {\n\n");
                int tabIndex = 0;
                AddTab(stringBuilder, ref tabIndex);
                stringBuilder.Append($"public static class {xlsxGenerateData.XlsxName}{XlsxDefine.CLASS_SUFFIX} " + "{\n\n");
                string structName = xlsxGenerateData.XlsxName + XlsxDefine.STRUCT_SUFFIX;
                stringBuilder.Append(CreateStruct(structName, xlsxGenerateData.HeadInfoDic, ref tabIndex));
                stringBuilder.Append(CreateDictionary(xlsxGenerateData, structName, ref tabIndex));
                stringBuilder.Append(CreateFunctions(xlsxGenerateData.HeadInfoDic, structName, ref tabIndex));
                stringBuilder.Append(XlsxDefine.TAB_CHAR_ARRAY[1] + "}\n\n");
                stringBuilder.Append("}\n");
                File.WriteAllText(xlsxGenerateData.OutputFilePath, stringBuilder.ToString());
                Log.Info($"{xlsxGenerateData.XlsxName}{XlsxDefine.CLASS_SUFFIX}.cs generate success.");
            } catch(Exception e) {
                throw e;
            }
        }

        private static string CreateStruct(string structName, Dictionary<int, XlsxHeadInfo> headInfoDic, ref int tabIndex) {
            StringBuilder stringBuilder = new StringBuilder();
            AddTab(stringBuilder, ref tabIndex);
            stringBuilder.Append($"public struct {structName} " + "{\n");
            tabIndex++;
            foreach(var value in headInfoDic.Values) {
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.Append($"public {XlsxFieldType.GetValueTypeString(value.ValueType)} {value.Name};\n");
            }
            SubTab(stringBuilder, ref tabIndex);
            stringBuilder.Append("}\n\n");
            return stringBuilder.ToString();
        }

        private static string CreateDictionary(XlsxGenerateData xlsxGenerateData, string structName, ref int tabIndex) {
            try {
                StringBuilder stringBuilder = new StringBuilder();
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.Append($"private readonly static Dictionary<int, {structName}> m_dic = new Dictionary<int, {structName}> " + "{\n");
                tabIndex++;
                foreach(var value in xlsxGenerateData.IdInfoDic.Values) {
                    KeepTap(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("{" + $" {value.Value}, new {structName} " + "{");
                    tabIndex++;
                    foreach(var headInfo in xlsxGenerateData.HeadInfoDic.Values) {
                        KeepTap(stringBuilder, ref tabIndex);
                        stringBuilder.AppendLine($"{headInfo.Name} = {GetValue(value, headInfo, xlsxGenerateData.Content, ref tabIndex)},");
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

        private static string GetValue(XlsxIDInfo idInfo, XlsxHeadInfo headInfo, object[,] content, ref int tabIndex) {
            try {
                object obj = content[idInfo.RowIndex, headInfo.Index];
                switch(headInfo.ValueType) {
                    case XlsxFieldType.STRING:
                        return TryConvertString(obj);
                    case XlsxFieldType.STRING_ARRAY:
                        return TryConvertStringArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.STRING_2D_ARRAY:
                        return TryConvertString2DArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);;
                    case XlsxFieldType.INT:
                        return TryConvertInt(obj);
                    case XlsxFieldType.INT_ARRAY:
                        return TryConvertIntArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.INT_2D_ARRAY:
                        return TryConvertInt2DArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.INT_TEN:
                        return TryConvertIntTen(obj);
                    case XlsxFieldType.INT_TEN_ARRAY:
                        return TryConvertIntTenArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.INT_TEN_2D_ARRAY:
                        return TryConvertIntTen2DArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.INT_HUNDRED:
                        return TryConvertIntHundred(obj);
                    case XlsxFieldType.INT_HUNDRED_ARRAY:
                        return TryConvertIntHundredArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.INT_HUNDRED_2D_ARRAY:
                        return TryConvertIntHundred2DArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.INT_THOUSAND:
                        return TryConvertIntThousand(obj);
                    case XlsxFieldType.INT_THOUSAND_ARRAY:
                        return TryConvertIntThousandArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.INT_THOUSAND_2D_ARRAY:
                        return TryConvertIntThousand2DArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.BOOLEAN:
                        return TryConvertBoolean(obj);
                    case XlsxFieldType.BOOLEAN_ARRAY:
                        return TryConvertBooleanArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                    case XlsxFieldType.BOOLEAN_2D_ARRAY:
                        return TryConvertBoolean2DArray(idInfo.RowIndex, headInfo.Index, content, idInfo.Height, ref tabIndex);
                }
                return "null";
            } catch(Exception e) {
                throw new Exception(e.Message + $" in ({idInfo.RowIndex}, {headInfo.Index}). ");
            }
        }

        private static string TryConvertIntTenArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.FLOAT_ARRAY, ref tabIndex, TryConvertIntTen);
        }
        
        private static string TryConvertIntHundredArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.FLOAT_ARRAY, ref tabIndex, TryConvertIntHundred);
        }

        private static string TryConvertIntThousandArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.FLOAT_ARRAY, ref tabIndex, TryConvertIntThousand);
        }

        private static string TryConvertStringArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.STRING_ARRAY, ref tabIndex, TryConvertString);
        }

        private static string TryConvertIntArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.INT_ARRAY, ref tabIndex, TryConvertInt);
        }

        private static string TryConvertBooleanArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvertArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.BOOLEAN_ARRAY, ref tabIndex, TryConvertBoolean);
        }

        private static string TryConvertArray(int rowIndex, int colIndex, object[,] content, int maxHeight, string valueType, ref int tabIndex, Func<object, string> tryConvertHandler) {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"new {valueType} " + "{");
            bool hasContent = false;
            tabIndex++;
            for(int row = rowIndex; row < rowIndex+maxHeight; row++) {
                if(content[row, colIndex] == null) {
                    break;
                }
                hasContent = true;
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.Append(tryConvertHandler(content[row, colIndex]));
                stringBuilder.AppendLine(",");
            }
            SubTab(stringBuilder, ref tabIndex);
            stringBuilder.Append("}");
            if(hasContent) {
                return stringBuilder.ToString();
            }
            return "null";
        }

        private static string TryConvertString2DArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.STRING, ref tabIndex, TryConvertString);
        }
        
        private static string TryConvertInt2DArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.INT, ref tabIndex, TryConvertInt);
        }

        private static string TryConvertIntTen2DArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.FLOAT, ref tabIndex, TryConvertIntTen);
        }

        private static string TryConvertIntHundred2DArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.FLOAT, ref tabIndex, TryConvertIntHundred);
        }

        private static string TryConvertIntThousand2DArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.FLOAT, ref tabIndex, TryConvertIntThousand);
        }

        private static string TryConvertBoolean2DArray(int rowIndex, int colIndex, object[,] content, int maxHeight, ref int tabIndex) {
            return TryConvert2DArray(rowIndex, colIndex, content, maxHeight, XlsxFieldType.BOOLEAN, ref tabIndex, TryConvertBoolean);
        }
        
        private static string TryConvert2DArray(int rowIndex, int colIndex, object[,] content, int maxHeight, string valueType, ref int tabIndex, Func<object, string> tryConvertHandler) {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"new {valueType}[][] " + "{");
            bool hasContent = false;
            tabIndex++;
            for(int row = rowIndex; row < rowIndex + maxHeight; row++) {
                if(content[row, colIndex] == null) {
                    break;
                }
                hasContent = true;
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"new {valueType}[] " + "{");
                string text = content[row, colIndex].ToString();
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

        private static string CreateFunctions(Dictionary<int, XlsxHeadInfo> headInfoDic, string structName, ref int tabIndex) {
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
                if(headInfo.Name == XlsxDefine.ID_CHAR || headInfo.Name == XlsxDefine.VALUE_CHAR) {
                    continue;
                }

                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"public static {XlsxFieldType.GetValueTypeString(headInfo.ValueType)} {headInfo.Name}(int id) " + "{");
                AddTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine("if(Contain(id)) {");
                AddTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"return m_dic[id].{headInfo.Name};");
                SubTab(stringBuilder,ref tabIndex);
                stringBuilder.AppendLine("}");
                KeepTap(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine($"return {XlsxFieldType.GetDefaultValue(headInfo.ValueType)};");
                SubTab(stringBuilder, ref tabIndex);
                stringBuilder.AppendLine("}\n");

                string returnType = "";
                bool isArray = XlsxFieldType.IsArray(headInfo.ValueType, ref returnType);
                if(isArray) {
                    KeepTap(stringBuilder,ref tabIndex);
                    stringBuilder.AppendLine($"public static {returnType} {headInfo.Name}FromArray(int id, int index)" + "{");
                    AddTab(stringBuilder,ref tabIndex);
                    stringBuilder.AppendLine("if(Contain(id)) {");
                    AddTab(stringBuilder,ref tabIndex);
                    stringBuilder.AppendLine($"{XlsxFieldType.GetValueTypeString(headInfo.ValueType)} array = {headInfo.Name}(id);");
                    KeepTap(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("if(array.Length > index) {");
                    AddTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("return array[index];");
                    SubTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("}");
                    SubTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("}");
                    KeepTap(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine($"return {XlsxFieldType.GetDefaultValue(returnType)};");
                    SubTab(stringBuilder, ref tabIndex);
                    stringBuilder.AppendLine("}\n");
                }

            }
            return stringBuilder.ToString();
        }

    }
}