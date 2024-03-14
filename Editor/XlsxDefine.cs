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
    }
}