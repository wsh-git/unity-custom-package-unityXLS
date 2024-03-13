
using UnityEditor;

namespace Wsh.XLS.Editor {

    public class XLSMenu {

        [MenuItem("Wsh/Xls/Builder", priority = 1)]
        public static void XlsEditorWindow() {
            XLSBuilderWindow.ShowWindow();
        }

        

    }

}

