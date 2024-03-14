using UnityEditor;

namespace Wsh.Xlsx.Editor {

    public class XlsxMenu {

        [MenuItem("Wsh/Xlsx/Builder", priority = 1)]
        public static void XlsxEditorWindow() {
            XlsxBuilderWindow.ShowWindow();
        }

    }

}

