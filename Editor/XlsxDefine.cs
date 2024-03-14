namespace Wsh.Xlsx.Editor {

    public static class XlsxDefine {

        public const string ID_CHAR = "ID";
        public const string VALUE_CHAR = "VALUE";
        public const string NULL_CHAR = "NULL";

        public const string STRUCT_SUFFIX = "XlsxData";
        public const string CLASS_SUFFIX = "XlsxWrapper";

        public const int XLSX_NULL_ID = 0;

        public readonly static string[] TAB_CHAR_ARRAY = new string[10]{
            "",
            "    ",
            "        ",
            "            ",
            "                ",
            "                    ",
            "                        ",
            "                            ",
            "                                ",
            "                                    ",
        };

        public const string XLS_EXTENSION = ".xls";
        public const string XLSX_EXTENSION = ".xlsx";

        public const string XLSX_ID_CLASS_NAME = "XlsxId";
        public const string XLSX_ID_FILE_NAME = XLSX_ID_CLASS_NAME + ".cs";

        public const string LOCAL_ID_CLASS_NAME = "LocalId";
        public const string LOCAL_ID_FILE_NAME = LOCAL_ID_CLASS_NAME + ".cs";
    }
}