using System.Collections.Generic;
using System.Text;

namespace Wsh.XLS.Editor {
    public class XLSFieldType {

        public const string STRING = "string";
        public const string STRING_ARRAY = "string[]";
        public const string STRING_2D_ARRAY = "string[][]";

        public const string INT = "int";
        public const string INT_ARRAY = "int[]";
        public const string INT_2D_ARRAY = "int[][]";

        public const string INT_TEN = "int10";
        public const string INT_TEN_ARRAY = "int10[]";
        public const string INT_TEN_2D_ARRAY = "int10[][]";

        public const string INT_HUNDRED = "int100";
        public const string INT_HUNDRED_ARRAY = "int100[]";
        public const string INT_HUNDRED_2D_ARRAY = "int100[][]";

        public const string INT_THOUSAND = "int1000";
        public const string INT_THOUSAND_ARRAY = "int1000[]";
        public const string INT_THOUSAND_2D_ARRAY = "int1000[][]";

        public const string BOOLEAN = "bool";
        public const string BOOLEAN_ARRAY = "bool[]";
        public const string BOOLEAN_2D_ARRAY = "bool[][]";

        public const string LOCAL = "local";
        public const string LOCAL_ARRAY = "local[]";
        public const string LOCAL_2D_ARRAY = "local[][]";

        public const string ID = "id";
        public const string ID_ARRAY = "id[]";
        public const string ID_2D_ARRAY = "id[][]";

        // float 不用在配置表中，为了程序使用
        public const string FLOAT = "float";
        public const string FLOAT_ARRAY = "float[]";
        public const string FLOAT_2D_ARRAY = "float[][]";

        private readonly static Dictionary<string, int> m_dic = new Dictionary<string, int> {
            {STRING, 1},
            {STRING_ARRAY, 1},
            {STRING_2D_ARRAY, 1},
            {INT, 1},
            {INT_ARRAY, 1},
            {INT_2D_ARRAY, 1},
            {INT_TEN, 1},
            {INT_TEN_ARRAY, 1},
            {INT_TEN_2D_ARRAY, 1},
            {INT_HUNDRED, 1},
            {INT_HUNDRED_ARRAY, 1},
            {INT_HUNDRED_2D_ARRAY, 1},
            {INT_THOUSAND, 1},
            {INT_THOUSAND_ARRAY, 1},
            {INT_THOUSAND_2D_ARRAY, 1},
            {BOOLEAN, 1},
            {BOOLEAN_ARRAY, 1},
            {BOOLEAN_2D_ARRAY, 1},
            {LOCAL, 1},
            {LOCAL_ARRAY, 1},
            {LOCAL_2D_ARRAY, 1},
            {ID, 1},
            {ID_ARRAY, 1},
            {ID_2D_ARRAY, 1},
        };

        public static string GetValueTypeString(string valueType) {
            if(valueType == INT_TEN || valueType == INT_HUNDRED || valueType == INT_THOUSAND) {
                return "float";
            } else if(valueType == INT_TEN_ARRAY || valueType == INT_HUNDRED_ARRAY || valueType == INT_THOUSAND_ARRAY) {
                return "float[]";
            } else if(valueType == INT_TEN_2D_ARRAY || valueType == INT_HUNDRED_2D_ARRAY || valueType == INT_THOUSAND_2D_ARRAY) {
                return "float[][]";
            } else if(valueType == LOCAL || valueType == ID) {
                return INT;
            } else if(valueType == LOCAL_ARRAY || valueType == ID_ARRAY) {
                return INT_ARRAY;
            } else if(valueType == LOCAL_2D_ARRAY || valueType == ID_2D_ARRAY) {
                return INT_2D_ARRAY;
            }
            return valueType;
        }

        public static bool IsArray(string valueType, ref string returnType) {
            switch(valueType) {
                case STRING_ARRAY:
                    returnType = STRING;
                    return true;
                case INT_ARRAY:
                case LOCAL_ARRAY:
                case ID_ARRAY:
                    returnType = INT;
                    return true;
                case INT_TEN_ARRAY:
                case INT_HUNDRED_ARRAY:
                case INT_THOUSAND_ARRAY:
                    returnType = FLOAT;
                    return true;
                case BOOLEAN_ARRAY:
                    returnType = BOOLEAN;
                    return true;
                case STRING_2D_ARRAY:
                    returnType = STRING_ARRAY;
                    return true;
                case INT_2D_ARRAY:
                case LOCAL_2D_ARRAY:
                case ID_2D_ARRAY:
                    returnType = INT_ARRAY;
                    return true;
                case INT_TEN_2D_ARRAY:
                case INT_HUNDRED_2D_ARRAY:
                case INT_THOUSAND_2D_ARRAY:
                    returnType = FLOAT_ARRAY;
                    return true;
                case BOOLEAN_2D_ARRAY:
                    returnType = BOOLEAN_ARRAY;
                    return true;
                default:
                    return false;
            }
        }

        public static bool Contain(string valueType) {
            return m_dic.ContainsKey(valueType); 
        }

        public static string GetAllDefineType() {
            StringBuilder sb = new StringBuilder();
            sb.Append("\nDefine:\n");
            foreach(var key in m_dic.Keys) {
                sb.Append("    ");
                sb.Append(key);
                sb.Append('\n');
            }
            return sb.ToString();
        }

        public static string GetDefaultValue(string valueType) {
            switch(valueType) {
                case STRING:
                    return "\"\"";
                case INT:
                case INT_TEN:
                case INT_HUNDRED:
                case INT_THOUSAND:
                case FLOAT:
                    return "0";
                case BOOLEAN:
                    return "false";
                case LOCAL:
                case ID:
                    return XLSDefine.XLS_NULL_ID.ToString();
                default:
                    return "null";
            }
        }
        
    }
}