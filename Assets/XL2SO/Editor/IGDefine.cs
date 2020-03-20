using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.Reflection;
using NPOI.SS.UserModel;

namespace XL2SO
{
    namespace IG
    {
        [System.Serializable]
        public class RowInfo
        {
            public string       InstanceName   = null;               // [GUI] Instance name for this row
            public List<string> Value          = new List<string>(); // [GUI] Cells value in this row
            public bool         IsListItemOpen = false;              // [GUI] Foldout state
        }

        [System.Serializable]
        public class General
        {
            public List<RowInfo> RowInfoList    = null;         // RowInfo list for selected cells 
            public FieldInfo[]   FieldInfoArray = null;         // [GUI] FieldInfo array extracted from target script 
            public Vector2       ScrollPosition = Vector2.zero; // [GUI] Current scroll position
            public GUIStyle      FoldoutStyle   = null;         // [GUI] GUI style for Foldout
        }

        [System.Serializable]
        public class ScriptPart
        {
            public bool       IsOpen     = true;   // [GUI] Foldout state
            public MonoScript SOSAsset   = null;   // [GUI] Input SOS (ScriptabelObject Script) Asset
            public bool       IsStartup  = true;   // SOSAsset has never been changed or not
            public bool       IsSOSAsset = false;  // SOSAsset is a subclass derived from ScriptableObject or not
            public bool       IsValidSOS = false;  // SOSAsset contains at least 1 field or not
        }

        [System.Serializable]
        public class ExcelPart
        {
            public IWorkbook    Book               = null;               // Excel Book data read from ExcelAsset
            public ISheet       Sheet              = null;               // Selected sheet data read by NPOI function
            public Vector2Int   StartCell          = Vector2Int.zero;    // [GUI] Start cell index in selected cells
            public Vector2Int   LastCell           = Vector2Int.zero;    // [GUI] Last cell index in selected cells
            public List<string> SheetNameList      = new List<string>(); // [GUI] Name list of valid sheets in ExcelAsset
            public ExcelViewer  EVWindow           = null;               // [GUI] ExcelViwer Reference
            public bool         IsOpen             = true;               // [GUI] Foldout state for title
            public Object       ExcelAsset         = null;               // [GUI] Input Excel Asset
            public string       ExcelPath          = string.Empty;       // [GUI] Excel Asset Path
            public int          SelectedSheetIndex = 0;                  // [GUI] Target sheet index in SheetNameList
            public bool         IsStartup          = true;               // ExcelAsset has never been changed or not
            public bool         IsExcelAsset       = false;              // ExcelAsset format is correct (.xls/.xlsx) or not
            public bool         IsValidExcel       = false;              // ExcelAsset contains at least 1 sheet which contains more than 2 rows
        }

        [System.Serializable]
        public class InstanceListPart
        {
            public enum NamingRules { BaseNameWithIndex, FieldValue };
            public Vector2      ScrollPosition      = Vector2.zero;                  // [GUI] Current Scroll Position
            public bool         IsOpen              = true;                          // [GUI] Foldout state
            public NamingRules  NamingRule          = NamingRules.BaseNameWithIndex; // [GUI] 0: [BaseName], [BaseName](1), [BaseName](2),... 
                                                                                     //       1: Each value of a specified field
            public string       BaseName            = string.Empty;                  // [GUI] Base name for NamingRule = BaseNameWithIndex
            public int          ReferenceFieldIndex = 0;                             // [GUI] Field index for NamingRule = FieldValue
        }

        [System.Serializable]
        public class OutputPart
        {
            public bool   IsOpen         = true;         // [GUI] Foldout state
            public Object OutputDirAsset = null;         // [GUI] Output Directory Asset
            public string OutputDirPath  = string.Empty; // [GUI] Path of Output Directory Asset
            public bool   ReqReloadExcel = false;        // Request for reloading Excel file after instance generation 
        }
    }
}
