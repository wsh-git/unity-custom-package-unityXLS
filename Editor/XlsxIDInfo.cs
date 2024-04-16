﻿namespace Wsh.Xlsx.Editor {

    public class XlsxIDInfo {

        public string Id => m_id;
        public int Value => m_value;
        public int RowIndex => m_rowIndex;
        public int Height => m_height;
        
        private string m_id;
        private int m_value;
        private int m_rowIndex;
        private int m_height; //在xlsx中，一个 id 下的属性内容最长有多少行，特别是有数组的情况下，没有数组的情况下一般为 1 行；


        public XlsxIDInfo(string id, int value, int rowIndex, int height) {
            m_id = id;
            m_value = value;
            m_rowIndex = rowIndex;
            m_height = height;
        }

        public void SetHeight(int height) {
            m_height = height;
        }
    
        public override string ToString() {
            return "{Id: " + Id + " Value: " + Value + " RowIndex: " + RowIndex + " Height: " + Height + "}";
        }

        public void Dispose() {
            m_id = null;
            m_value = 0;
            m_rowIndex = 0;
            m_height = 0;
        }

    }

}