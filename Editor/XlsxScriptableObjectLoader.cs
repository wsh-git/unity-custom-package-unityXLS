using System;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace Wsh.Xlsx.Editor {

    public class XlsxScriptableObjectLoader {

        private const string SCRIPTABLEOBJECT_FOLDER = "Assets/WshConfig/";

        private const string SCRIPTABLEOBJECT_PATH = SCRIPTABLEOBJECT_FOLDER + "XlsxMainWindowScriptableObject.asset";

        public XlsxMainScriptableObject LoadXlsMainScriptableObject() {
            XlsxMainScriptableObject scriptableObject = AssetDatabase.LoadAssetAtPath<XlsxMainScriptableObject>(SCRIPTABLEOBJECT_PATH);
            return scriptableObject;
        }
        
        public void SaveScriptableObject(XlsxBuilderWindow window) {
            XlsxMainScriptableObject scriptableObject = AssetDatabase.LoadAssetAtPath<XlsxMainScriptableObject>(SCRIPTABLEOBJECT_PATH);
            if(scriptableObject == null) {
                scriptableObject = ScriptableObject.CreateInstance<XlsxMainScriptableObject>();
                AssetDatabase.CreateAsset(scriptableObject, SCRIPTABLEOBJECT_PATH);
            }
            scriptableObject.XlsxDir = window.XlsxDir;
            scriptableObject.XlsxDataOutputDir = window.XlsxDataOutputDir;
            EditorUtility.SetDirty(scriptableObject);
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }

        public void CheckScriptableObject() {
            Type scriptableObj = typeof(XlsxMainScriptableObject);
            string[] assetPaths = AssetDatabase.FindAssets("t:" + scriptableObj);
            if(assetPaths == null || assetPaths.Length == 0) {
                TryCreateScriptableObject();
            }
        }

        private void TryCreateScriptableObject() {
            XlsxMainScriptableObject scriptableObject = ScriptableObject.CreateInstance<XlsxMainScriptableObject>();
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