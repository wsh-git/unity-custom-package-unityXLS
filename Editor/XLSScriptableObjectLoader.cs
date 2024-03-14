using System;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace Wsh.XLS.Editor {

    public class XLSScriptableObjectLoader {

        private const string SCRIPTABLEOBJECT_FOLDER = "Assets/WshConfig/";

        private const string SCRIPTABLEOBJECT_PATH = SCRIPTABLEOBJECT_FOLDER + "XLSMainWindowScriptableObject.asset";

        public XLSMainScriptableObject LoadXlsMainScriptableObject() {
            XLSMainScriptableObject scriptableObject = AssetDatabase.LoadAssetAtPath<XLSMainScriptableObject>(SCRIPTABLEOBJECT_PATH);
            return scriptableObject;
        }
        
        public void SaveScriptableObject(XLSBuilderWindow window) {
            XLSMainScriptableObject scriptableObject = AssetDatabase.LoadAssetAtPath<XLSMainScriptableObject>(SCRIPTABLEOBJECT_PATH);
            if(scriptableObject == null) {
                scriptableObject = ScriptableObject.CreateInstance<XLSMainScriptableObject>();
                AssetDatabase.CreateAsset(scriptableObject, SCRIPTABLEOBJECT_PATH);
            }
            scriptableObject.XlsDir = window.XlsDir;
            scriptableObject.XlsDataOutputDir = window.XlsDataOutputDir;
            EditorUtility.SetDirty(scriptableObject);
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }

        public void CheckScriptableObject() {
            Type scriptableObj = typeof(XLSMainScriptableObject);
            string[] assetPaths = AssetDatabase.FindAssets("t:" + scriptableObj);
            if(assetPaths == null || assetPaths.Length == 0) {
                TryCreateScriptableObject();
            }
        }

        private void TryCreateScriptableObject() {
            XLSMainScriptableObject scriptableObject = ScriptableObject.CreateInstance<XLSMainScriptableObject>();
            TryCreateFolder();
            AssetDatabase.CreateAsset(scriptableObject, SCRIPTABLEOBJECT_PATH);
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }

        private void TryCreateFolder() {
            if(!Directory.Exists(SCRIPTABLEOBJECT_FOLDER)) {
                Directory.CreateDirectory(SCRIPTABLEOBJECT_FOLDER);
            }
        }

    }
}