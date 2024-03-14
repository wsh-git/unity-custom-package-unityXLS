using System.Security.Principal;
using UnityEditor;
using UnityEngine;

namespace Wsh.XLS.Editor {

    public class XLSBuilderWindow : EditorWindow {

        private const int FIRST_SPACE = 15;
        private const string EMPTY_STRING = "<None>";

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

        public string XlsDir => m_xlsDir;
        public string XlsDataOutputDir => m_xlsDataOutputDir;
        
        private string m_xlsDir;
        private string m_xlsDataOutputDir;
        private XLSScriptableObjectLoader m_scriptableObjLoader;
        private bool m_isInited;
        
        private void InitMainWindowData() {
            var data = m_scriptableObjLoader.LoadXlsMainScriptableObject();
            if(data != null) {
                m_xlsDir = data.XlsDir;
                m_xlsDataOutputDir = data.XlsDataOutputDir;
                m_isInited = true;
            }
        }
        
        private void OnEnable() {
            m_scriptableObjLoader = new XLSScriptableObjectLoader();
            m_scriptableObjLoader.CheckScriptableObject();
            InitMainWindowData();
        }
        
        private void OnDisable() {
            m_instance = null;
        }

        private void OnGUI() {
            if(!m_isInited) { return; }
            GUILayout.Space(10);
            #region ResRootDir
            GUILayout.BeginHorizontal();
            GUILayout.Space(FIRST_SPACE);
            GUILayout.Label("XlsDir:", GUILayout.Width(100));
            if(string.IsNullOrEmpty(m_xlsDir)) {
                GUILayout.Label(EMPTY_STRING);
            } else {
                GUILayout.Label(m_xlsDir);
            }
            if(GUILayout.Button("浏览")) {
                string path = EditorUtility.OpenFolderPanel("选择项目Xls路径", Application.dataPath, "");
                if(path != null) {
                    m_xlsDir = path;
                }
            }
            GUILayout.EndHorizontal();
            #endregion

            #region OutputDir
            GUILayout.BeginHorizontal();
            GUILayout.Space(FIRST_SPACE);
            GUILayout.Label("OutputDir:", GUILayout.Width(100));
            if(string.IsNullOrEmpty(m_xlsDataOutputDir)) {
                GUILayout.Label(EMPTY_STRING);
            } else {
                GUILayout.Label(m_xlsDataOutputDir);
            }
            if(GUILayout.Button("浏览")) {
                string path = EditorUtility.OpenFolderPanel("选择Xls数据导出路径", Application.dataPath, "");
                if(path != null) {
                    m_xlsDataOutputDir = path;
                }
            }
            GUILayout.EndHorizontal();
            #endregion

            GUILayout.Space(10);

            #region Save Config
            GUILayout.BeginHorizontal();
            if(GUILayout.Button("Save Config", GUILayout.Height(30))) {
                m_scriptableObjLoader.SaveScriptableObject(this);
            }
            GUILayout.EndHorizontal();
            #endregion

            GUILayout.Space(10);
            if(GUILayout.Button("Build All Xls", GUILayout.Height(30))) {
                XLSBuilder.BuildFolder(m_xlsDir, m_xlsDataOutputDir);
            }

            GUILayout.Space(10);
            if(GUILayout.Button("Clear All Generate", GUILayout.Height(30))) {
                XLSBuilder.ClearFolder(m_xlsDataOutputDir);
            }
        }

    }

}
