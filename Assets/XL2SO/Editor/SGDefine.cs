using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using NPOI.SS.UserModel;
using System.Reflection;

namespace XL2SO
{
    namespace SG
    {
        public enum PrimaryDataTypes { @char, @string, @sbyte, @byte, @short, @ushort, @int, @uint, @long, @ulong, @float, @double, @decimal, @bool};
        public enum AccessLevels     { @public, @protected, @private };

        [System.Serializable]
        public class ColumnInfo
        {
            public bool         Ignore = false;                // [GUI] Indicate that this field is ignored
            public AccessLevels Access = AccessLevels.@public; // [GUI] Accessibility level
            public string       Type   = string.Empty;         // [GUI] Data type
            public List<string> Value  = new List<string>();   // [GUI] Cells value in this column
                                                               //       Value[0] corresponds to column label
        }

        [System.Serializable]
        public class General
        {
            public string           TemplatePath    = string.Empty;            // Path of template file
            public bool             TemplateExists  = false;                   // Template file exists in TemplatePath or not
            public List<ColumnInfo> ColumnInfoList  = new List<ColumnInfo>();  // ColumnInfo list for selected cells
            public Vector2          ScrollPosition  = Vector2.zero;            // [GUI] Current scroll position
            public GUIStyle         FoldoutStyle    = null;                    // [GUI] GUIStyle for Foldout element 
            public const string     CharPrefix      = "Character/";            // Prefix for grouping char type 
            public const string     IntPrefix       = "Integer/";              // Prefix for grouping integer type 
            public const string     RealPrefix      = "Real/";                 // Prefix for grouping real type 
            public const string     EnumPrefix      = "Enum/";                 // Prefix for grouping enum type 
        }

        [System.Serializable]
        public class ExcelPart
        {
            public IWorkbook    Book               = null;               // Excel Book data read from ExcelAsset
            public ISheet       Sheet              = null;               // Selected sheet data read from Book
            public Vector2Int   StartCell          = Vector2Int.zero;    // [GUI] Start cell index in selected cells
            public Vector2Int   LastCell           = Vector2Int.zero;    // [GUI] Last cell index in selected cells
            public List<string> SheetNameList      = new List<string>(); // [GUI] Name list of valid sheets in ExcelAsset
            public ExcelViewer  EVWindow           = null;               // [GUI] ExcelViwer instance reference 
            public bool         IsOpen             = true;               // [GUI] Foldout state
            public Object       ExcelAsset         = null;               // [GUI] Input Excel Asset
            public string       ExcelPath          = string.Empty;       // [GUI] Excel Asset Path
            public int          SelectedSheetIndex = 0;                  // [GUI] Target sheet index in SheetNameList
            public bool         IsStartup          = true;               // ExcelAsset has never been changed or not
            public bool         IsExcelAsset       = false;              // ExcelAsset format is correct (.xls/.xlsx) or not
            public bool         IsValidExcel       = false;              // ExcelAsset contains at least 1 sheet which contains more than 1 rows
        }

        [System.Serializable]
        public class FieldListPart
        {
            public Vector2      ScrollPosition = Vector2.zero;           // [GUI] Current Scroll Position
            public bool         IsOpen         = true;                   // [GUI] Foldout state
            public List<string> DataTypeList   = null;                   // [GUI] Data type list including primary data type and user-defined enum
            public GUIStyle     ColumnStyle    = null;                   // [GUI] GUIStyle for table column 
            public GUIStyle     ContentStyle   = null;                   // [GUI] GUIStyle for table contents 
            public GUIStyle     PopupStyle     = null;                   // [GUI] GUIStyle for Popup element
            public bool         IgnoreAll      = false;                  // [GUI] Ignore setting for all items
            public string       TypeAll        = string.Empty;           // [GUI] Data type setting for all items
            public AccessLevels AccessAll      = AccessLevels.@public;   // [GUI] Accessibility setting for all items
        }

        [System.Serializable]
        public class OutputPart
        {
            public bool   IsOpen                         = true;         // [GUI] Foldout state
            public Object OutputDirAsset                 = null;         // [GUI] Output Directory Asset
            public string OutputDirPath                  = string.Empty; // Path of Output Directory Asset
            public string OutputFileNameWithoutExtension = string.Empty; // [GUI] Output File Name without extension
            public bool   ReqReloadExcel                 = false;        // Request for reloading Excel file after script generation. 
        }
    }
}
