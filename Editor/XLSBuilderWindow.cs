using UnityEditor;
using UnityEngine;

namespace Wsh.XLS.Editor {

    public class XLSBuilderWindow : EditorWindow {

        private static XLSBuilderWindow m_instance = null;

        public static XLSBuilderWindow Instance {
            get {
                if(m_instance == null) {
                    m_instance = GetWindow<XLSBuilderWindow>();
                }
                return m_instance;
            }
        }

        public static void ShowWindow() {
            m_instance = null;
            Instance.titleContent = new GUIContent("Xls Builder");
            Instance.Show();
        }

        private void OnDisable() {
            m_instance = null;
        }

        private void OnGUI() {
            GUILayout.Space(10);
            if(GUILayout.Button("Build", GUILayout.Height(30))) {
                XLSBuilder.BuildFile("Test", "D:/Projects/yiyiyaya/Excel/Test.xlsx", "D:/Projects/yiyiyaya/Excel/TestXlsWrapper.cs");
            }
        }

    }

}
