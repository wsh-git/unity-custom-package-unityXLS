using System.Security.Principal;
using UnityEditor;
using UnityEngine;

namespace Wsh.Xlsx.Editor {

    public class XlsxBuilderWindow : EditorWindow {

        private const int FIRST_SPACE = 15;
        private const string EMPTY_STRING = "<None>";

        private static XlsxBuilderWindow m_instance = null;

        public static XlsxBuilderWindow Instance {
            get {
                if(m_instance == null) {
                    m_instance = GetWindow<XlsxBuilderWindow>();
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
        private XlsxScriptableObjectLoader m_scriptableObjLoader;
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
            m_scriptableObjLoader = new XlsxScriptableObjectLoader();
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
            GUILayout.Label("XlsxDir:", GUILayout.Width(100));
            if(string.IsNullOrEmpty(m_xlsDir)) {
                GUILayout.Label(EMPTY_STRING);
            } else {
                GUILayout.Label(m_xlsDir);
            }
            if(GUILayout.Button("浏览")) {
                string path = EditorUtility.OpenFolderPanel("选择项目Xlsx路径", Application.dataPath, "");
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
                string path = EditorUtility.OpenFolderPanel("选择Xlsx数据导出路径", Application.dataPath, "");
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
            if(GUILayout.Button("Build All Xlsx", GUILayout.Height(30))) {
                XlsxBuilder.BuildFolder(m_xlsDir, m_xlsDataOutputDir);
            }

            GUILayout.Space(10);
            if(GUILayout.Button("Clear All Generate", GUILayout.Height(30))) {
                XlsxBuilder.ClearFolder(m_xlsDataOutputDir);
            }
        }

    }

}
