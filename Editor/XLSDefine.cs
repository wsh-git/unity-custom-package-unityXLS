namespace Wsh.XLS.Editor {

    public static class XLSDefine {

        public const string ID_CHAR = "ID";
        public const string VALUE_CHAR = "VALUE";
        public const string NULL_CHAR = "NULL";

        public const string STRUCT_SUFFIX = "XlsData";
        public const string CLASS_SUFFIX = "XlsWrapper";

        public const int XLS_NULL_ID = 0;

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
    }
}