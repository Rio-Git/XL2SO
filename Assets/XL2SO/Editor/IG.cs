using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.Linq;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Reflection;

namespace XL2SO
{
    namespace IG
    {
        /// <summary>
        /// Represents a "Generate ScriptableObject Instance" feature.
        /// </summary>
        public class IG : IScreen
        {
            private XL2SO            m_Parent           = null;
            private General          m_General          = null;
            private ScriptPart       m_ScriptPart     = null;
            private ExcelPart        m_ExcelPart        = null;
            private InstanceListPart m_InstanceListPart = null;
            private OutputPart       m_OutputPart       = null;

            /// <summary>
            /// Initializes a new instance of the <see cref="IG"/> class.
            /// <summary>
            /// <param name="_Parent"></param>
            override public void Initialize(XL2SO _Parent)
            {
                m_Parent = _Parent;

                m_General          = new General();
                m_ScriptPart       = new ScriptPart();
                m_ExcelPart        = new ExcelPart();
                m_InstanceListPart = new InstanceListPart();
                m_OutputPart       = new OutputPart();
                
                m_General.FoldoutStyle = new GUIStyle("FoldOut");
                m_General.FoldoutStyle.fontStyle = FontStyle.Bold;
            }

            /// <summary>
            /// Main routine of <see cref="IG"/> package.
            /// </summary>
            /// <returns>
            /// A new instance of <see cref="Menu"/> class: When "Return to Menu" button is clicked.
            /// A instance of <see cref="IG"/> class (this): Other than above case.
            /// </returns>
            override public IScreen Display()
            {
                m_General.ScrollPosition = EditorGUILayout.BeginScrollView(m_General.ScrollPosition);
                if (ShowScriptPart()) {
                    if (ShowExcelPart()) {
                        if (ShowInstanceListPart())
                            ShowOutputPart();
                    }
                }
                EditorGUILayout.EndScrollView();

                if (ShowMenuButton())
                    return CreateInstance<Menu>();
                else
                    return this;
            }

            /// <summary>
            /// Shows ScriptableObejct Script Part.
            /// </summary>
            /// <remarks>
            /// Error checkers:
            /// - Format check: The specified script is a sub-class of ScriptableObject or not.
            /// - Contents check: The specified script contains at least 1 field or not.
            /// 
            /// Flag Combination Table: 
            /// - IsStatup | IsValidSOS | IsSOSAsset   
            /// -     1    |       0    |       0     => SOS (ScriptableObject Script) asset has never changed.
            /// -     0    |       0    |       0     => Format error.
            /// -     0    |       0    |       1     => Contents error.
            /// -     0    |       1    |       1     => Complete all process correctly
            /// 
            /// </remarks>
            /// <returns>
            /// - True: No error occurred.
            /// - False: Error occurred.
            /// </returns>
            private bool ShowScriptPart()
            {
                m_ScriptPart.IsOpen = EditorGUILayout.Foldout(m_ScriptPart.IsOpen, new GUIContent("ScriptableObject Script"), m_General.FoldoutStyle);
                if (m_ScriptPart.IsOpen) {
                    EditorGUILayout.BeginHorizontal();
                    {
                        EditorGUILayout.LabelField("Script", GUILayout.MaxWidth(50));

                        GUILayout.FlexibleSpace();

                        EditorGUI.BeginChangeCheck();
                        m_ScriptPart.SOSAsset = EditorGUILayout.ObjectField(m_ScriptPart.SOSAsset, typeof(MonoScript), true, GUILayout.Width(150)) as MonoScript;
                        if (EditorGUI.EndChangeCheck()) {
                            m_ScriptPart.IsStartup = false;
                            
                            // Format check
                            Type type = TryGetDividedType(m_ScriptPart.SOSAsset.name, typeof(ScriptableObject));
                            if (type != null) {
                                m_ScriptPart.IsSOSAsset = true;
                                
                                // Contents check
                                m_General.FieldInfoArray = type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                                if (m_General.FieldInfoArray.Length == 0)
                                    m_ScriptPart.IsValidSOS = false;
                                else {
                                    m_ScriptPart.IsValidSOS = true;
                                    string sos_asset_path = AssetDatabase.GetAssetPath(m_ScriptPart.SOSAsset);
                                    string sos_asset_dir_path = Path.GetDirectoryName(sos_asset_path);
                                    m_OutputPart.OutputDirAsset = AssetDatabase.LoadAssetAtPath(sos_asset_dir_path, typeof(UnityEngine.Object));
                                    m_OutputPart.OutputDirPath = sos_asset_dir_path;
                                    m_InstanceListPart.BaseName = m_ScriptPart.SOSAsset.name + "Instance";

                                    // If Excel is already read, reset RowIntoList with new FieldInfoArray.
                                    if (!m_ExcelPart.IsStartup & m_ExcelPart.IsExcelAsset & m_ExcelPart.IsValidExcel)
                                        InitializeRowInfoList();
                                }
                            } else {
                                m_ScriptPart.IsSOSAsset = false;
                                m_ScriptPart.IsValidSOS = false;
                            }
                        }
                    }
                    EditorGUILayout.EndHorizontal();
                }

                // Display Error Message
                if (!m_ScriptPart.IsStartup) {
                    if (!m_ScriptPart.IsSOSAsset)
                        EditorGUILayout.HelpBox("Please set ScriptableObject script.", MessageType.Error);
                    else if (!m_ScriptPart.IsValidSOS)
                        EditorGUILayout.HelpBox("Not Found any fields.", MessageType.Error);
                }

                return !m_ScriptPart.IsStartup & m_ScriptPart.IsSOSAsset & m_ScriptPart.IsValidSOS;
            }

            /// <summary>
            /// Try to get type of "_name" class which is a derived class of _base_type
            /// </summary>
            /// <param name="_name">Class name to be found.</param>
            /// <param name="_base_type">Base type of a class to be found.</param>
            /// <returns>
            /// Type of found class.
            /// If not found, returns null.
            /// </returns>
            public Type TryGetDividedType (string _name, Type _base_type)
            {
                Type result = null;
                foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies()) {
                    foreach (Type type in assembly.GetTypes()) {
                        if ((type.Name == _name) && type.BaseType == typeof(ScriptableObject))
                            result = type;
                    }
                }
                return result;
            }

            /// <summary>
            /// Shows Excel Part.
            /// </summary>
            /// <remarks>
            /// Error checkers:
            /// - Format check: The specified Excel asset is .xls or .xlsx file or not.
            /// - Contents check: The Excel has at least 1 sheet which contains at least 2 row (for column label and data) or not.
            /// 
            /// Flag Combination Table: 
            /// - IsStatup | IsValidExcel | IsExcelAsset   
            /// -     1    |       0      |       0      => Excel asset has never changed.
            /// -     0    |       0      |       0      => Format error.
            /// -     0    |       0      |       1      => Contents error.
            /// -     0    |       1      |       1      => Complete all process correctly.
            /// 
            /// Request for reloading Excel:
            /// - When generating instance is done by GenerateInstance(), Unity do serialization/deserializatoin unexpectedly and 
            /// non-serializable member like Book and Sheet are set to Null.
            /// - This situation can cause NullReferenceExeption when referring Book or Sheet in the following process.
            ///     > Change a sheet which refers Book
            ///     > Open Excel Viewer which refers Sheet
            /// - To avoid this, it needs to reload Excel and set Book and Sheet when m_OutputPart.ReqReloadExcel == true.
            /// </remarks>
            /// <returns>
            /// - True: No error occurred.
            /// - False: Error occurred.
            /// </returns>
            private bool ShowExcelPart()
            {
                m_ExcelPart.IsOpen = EditorGUILayout.Foldout(m_ExcelPart.IsOpen,
                                                             new GUIContent("Excel"), m_General.FoldoutStyle);
                if (m_ExcelPart.IsOpen) {
                    EditorGUI.BeginChangeCheck();
                    EditorGUILayout.BeginHorizontal();
                    {
                        EditorGUILayout.LabelField("Book", GUILayout.MaxWidth(50));
                        GUILayout.FlexibleSpace();
                        m_ExcelPart.ExcelAsset = EditorGUILayout.ObjectField(m_ExcelPart.ExcelAsset,
                                                                             typeof(UnityEngine.Object), false, GUILayout.Width(150));
                    }
                    EditorGUILayout.EndHorizontal();
                    if (EditorGUI.EndChangeCheck()) {
                        m_ExcelPart.IsStartup = false;

                        if (m_ExcelPart.EVWindow)
                            m_ExcelPart.EVWindow.Close();

                        // Format check
                        m_ExcelPart.ExcelPath = AssetDatabase.GetAssetPath(m_ExcelPart.ExcelAsset);
                        if ((Path.GetExtension(m_ExcelPart.ExcelPath) == ".xls") ||
                            (Path.GetExtension(m_ExcelPart.ExcelPath) == ".xlsx")) {
                            m_ExcelPart.IsExcelAsset = true;

                            FileStream fs = File.Open(m_ExcelPart.ExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            if (Path.GetExtension(m_ExcelPart.ExcelPath) == ".xls")
                                m_ExcelPart.Book = new HSSFWorkbook(fs);
                            else
                                m_ExcelPart.Book = new XSSFWorkbook(fs);
                            fs.Close();

                            // Contents check
                            m_ExcelPart.SheetNameList.Clear();
                            m_ExcelPart.SelectedSheetIndex = 0;
                            for (int sheet_idx = 0; sheet_idx < m_ExcelPart.Book.NumberOfSheets; sheet_idx++) {
                                ISheet sheet = m_ExcelPart.Book.GetSheetAt(sheet_idx);
                                if (sheet.PhysicalNumberOfRows > 1) {
                                    m_ExcelPart.SheetNameList.Add(sheet.SheetName);
                                }
                            }

                            if (m_ExcelPart.SheetNameList.Count > 0) {
                                m_ExcelPart.Sheet = m_ExcelPart.Book.GetSheet(m_ExcelPart.SheetNameList.First());
                                InitializeRowInfoList();
                                m_ExcelPart.IsValidExcel = true;
                            } else
                                m_ExcelPart.IsValidExcel = false;
                        } else {
                            m_ExcelPart.IsExcelAsset = false;
                            m_ExcelPart.IsValidExcel = false;
                        }
                    }

                    // Error Message
                    if (!m_ExcelPart.IsStartup) {
                        if (!m_ExcelPart.IsExcelAsset)
                            EditorGUILayout.HelpBox("Please set .xls or .xlsx file.", MessageType.Error);
                        else if (!m_ExcelPart.IsValidExcel)
                            EditorGUILayout.HelpBox("No valid sheets were found.", MessageType.Error);
                    }

                    if (m_ExcelPart.IsValidExcel) {
                        EditorGUILayout.BeginHorizontal();
                        {
                            EditorGUILayout.LabelField("Sheet");
                            EditorGUI.BeginChangeCheck();
                            m_ExcelPart.SelectedSheetIndex = EditorGUILayout.Popup(m_ExcelPart.SelectedSheetIndex,
                                                                                   m_ExcelPart.SheetNameList.ToArray(), GUILayout.Width(150));
                            if (EditorGUI.EndChangeCheck()) {
                                if (m_ExcelPart.EVWindow)
                                    m_ExcelPart.EVWindow.Close();
                                
                                // Reload Excel file after the script is generated 
                                if (m_OutputPart.ReqReloadExcel) {
                                    FileStream fs = File.Open(m_ExcelPart.ExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                                    if (Path.GetExtension(m_ExcelPart.ExcelPath) == ".xls")
                                        m_ExcelPart.Book = new HSSFWorkbook(fs);
                                    else
                                        m_ExcelPart.Book = new XSSFWorkbook(fs);
                                    fs.Close();

                                    m_OutputPart.ReqReloadExcel = false;
                                }

                                m_ExcelPart.Sheet = m_ExcelPart.Book.GetSheet(m_ExcelPart.SheetNameList[m_ExcelPart.SelectedSheetIndex]);
                                InitializeRowInfoList();

                            }
                        }
                        EditorGUILayout.EndHorizontal();

                        EditorGUILayout.BeginHorizontal();
                        {
                            EditorGUILayout.LabelField("Cells");
                            string format = "({0},{1}) to ({2},{3})";
                            int start_x   = Mathf.Min(m_ExcelPart.StartCell.x, m_ExcelPart.LastCell.x);
                            int start_y   = Mathf.Min(m_ExcelPart.StartCell.y, m_ExcelPart.LastCell.y);
                            int end_x     = Mathf.Max(m_ExcelPart.StartCell.x, m_ExcelPart.LastCell.x);
                            int end_y     = Mathf.Max(m_ExcelPart.StartCell.y, m_ExcelPart.LastCell.y);
                            string cells = string.Format(format, start_x, start_y, end_x, end_y);
                            EditorGUILayout.LabelField(cells, GUILayout.Width(150));
                        }
                        EditorGUILayout.EndHorizontal();

                        EditorGUILayout.BeginHorizontal();
                        {
                            GUILayout.FlexibleSpace();
                            if (GUILayout.Button("Select Cells", GUILayout.Width(150))) {
                                
                                // Reload Excel file after the instance is generated 
                                if (m_OutputPart.ReqReloadExcel) {
                                    FileStream fs = File.Open(m_ExcelPart.ExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                                    if (Path.GetExtension(m_ExcelPart.ExcelPath) == ".xls")
                                        m_ExcelPart.Book = new HSSFWorkbook(fs);
                                    else
                                        m_ExcelPart.Book = new XSSFWorkbook(fs);
                                    fs.Close();
                                    ISheet select_sheet = m_ExcelPart.Book.GetSheet(m_ExcelPart.SheetNameList[m_ExcelPart.SelectedSheetIndex]);
                                    m_ExcelPart.Sheet = select_sheet;
                                    m_OutputPart.ReqReloadExcel = false;
                                }

                                m_ExcelPart.EVWindow = EditorWindow.GetWindow<ExcelViewer>("ExcelViewer");
                                m_ExcelPart.EVWindow.Initialize(m_ExcelPart.StartCell,
                                                                m_ExcelPart.LastCell,
                                                                m_ExcelPart.Sheet,
                                                                OnSetSelectCells);
                            }
                        }
                        EditorGUILayout.EndHorizontal();
                    }

                }
                return !m_ExcelPart.IsStartup & m_ExcelPart.IsExcelAsset & m_ExcelPart.IsValidExcel;
            }
            
            /// <summary>
            /// Initializes RowInfoList with a default range.
            /// </summary>
            /// <remarks>
            /// Default range is set as below:
            /// - Column: FirstRow.FirstCellNum - FirstRow.LastCellNum
            /// - Row: FirstRowNum - LastRowNum
            /// This means first row of the sheet is used as a reference of column label.
            /// </remarks>
            private void InitializeRowInfoList()
            {
                IRow first_row    = m_ExcelPart.Sheet.GetRow(m_ExcelPart.Sheet.FirstRowNum);
                int start_col_idx = first_row.FirstCellNum;
                int start_row_idx = m_ExcelPart.Sheet.FirstRowNum;
                int last_col_idx  = first_row.LastCellNum - 1;
                int last_row_idx  = m_ExcelPart.Sheet.LastRowNum;
                UpdateRowInfoList(new Vector2Int(start_col_idx, start_row_idx), new Vector2Int(last_col_idx, last_row_idx));
            }

            /// <summary>
            /// Updates RowInfoList by a range of _start_cell and _start_cell.
            /// </summary>
            /// <remarks>
            /// RowInfo is added into RowInfoList only when: 
            /// - More than 2 rows are selected (one for column label and others for each instance data).
            /// - At least 1 column label matches with one of the field name in FieldInfoArray.
            /// </remarks>
            /// <param name="_start_cell"></param>
            /// <param name="_last_cell"></param>
            public void UpdateRowInfoList(Vector2Int _start_cell, Vector2Int _last_cell)
            {
                int start_col_idx = Mathf.Min(_start_cell.x, _last_cell.x);
                int last_col_idx  = Mathf.Max(_start_cell.x, _last_cell.x);
                int start_row_idx = Mathf.Min(_start_cell.y, _last_cell.y);
                int last_row_idx  = Mathf.Max(_start_cell.y, _last_cell.y);

                m_General.RowInfoList = new List<RowInfo>();

                // Nothing to do when only one row is selected
                if (start_row_idx == last_row_idx)
                    return;

                // Get first row which contain column labels
                IRow first_row = m_ExcelPart.Sheet.GetRow(start_row_idx);
                // Nothing to do when the start row is blank
                if (first_row == null)
                    return;

                foreach (FieldInfo field in m_General.FieldInfoArray) {
                    for (int col_idx = start_col_idx; col_idx <= last_col_idx; col_idx++) {
                        // Check if the current column label (first_row.GetCell(col_idx)) matches with current filed name
                        if (field.Name == first_row.GetCell(col_idx, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString()) {
                            // Reserve row info with empty
                            if (m_General.RowInfoList.Count == 0) {
                                for (int row_idx = start_row_idx; row_idx <= last_row_idx; row_idx++) {
                                    m_General.RowInfoList.Add(new RowInfo());
                                }
                            }

                            // Add cell value to each RowInfo.Values
                            // Instance data is added when the data can be casted to the field type
                            for (int row_idx = start_row_idx; row_idx <= last_row_idx; row_idx++) {
                                // 0-base index for RowInfoList and RowInfo.Value
                                int row_info_idx = row_idx - start_row_idx;
                                ICell cell = m_ExcelPart.Sheet.GetRow(row_idx).GetCell(col_idx);
                                if (cell != null) {
                                    if (row_idx == start_row_idx) // Column Label
                                        m_General.RowInfoList[row_info_idx].Value.Add(cell.ToString());
                                    else { // Instance data 
                                        // if the value can't be casted to the field type, set the value to empty
                                        if (TypeCast(field.FieldType, cell.ToString()) == null)
                                            m_General.RowInfoList[row_info_idx].Value.Add(string.Empty);
                                        else
                                            m_General.RowInfoList[row_info_idx].Value.Add(cell.ToString());
                                    }
                                } else // Blank cell
                                    m_General.RowInfoList[row_info_idx].Value.Add(string.Empty);
                            }
                        }
                    }
                }
                m_ExcelPart.StartCell = _start_cell;
                m_ExcelPart.LastCell  = _last_cell;
                SetInstanceName();
            }

            /// <summary>
            /// Set each instance name depending on NamingRule
            /// </summary>
            private void SetInstanceName()
            {
                List<string> exist_names = new List<string>();
                switch (m_InstanceListPart.NamingRule) {
                    case (InstanceListPart.NamingRules.BaseNameWithIndex):
                        for (int row_idx = 1; row_idx < m_General.RowInfoList.Count; row_idx++) {
                            string unique_name = ObjectNames.GetUniqueName(exist_names.ToArray(), m_InstanceListPart.BaseName);
                            m_General.RowInfoList[row_idx].InstanceName = unique_name;
                            exist_names.Add(unique_name);
                        }
                        break;
                    case (InstanceListPart.NamingRules.FieldValue):
                        string field_name = m_General.FieldInfoArray[m_InstanceListPart.ReferenceFieldIndex].Name;
                        int val_idx = m_General.RowInfoList[0].Value.IndexOf(field_name);
                        for (int row_idx = 1; row_idx < m_General.RowInfoList.Count; row_idx++) {
                            // Check if there is a RowInfo in RowInfoList which matches to field_name
                            string instance_name;
                            if (val_idx == -1)  // If it's not found, set instance name to empty
                                instance_name = string.Empty;
                            else  // If it's found, set instance name with its value
                                instance_name = m_General.RowInfoList[row_idx].Value[val_idx];

                            string unique_name = ObjectNames.GetUniqueName(exist_names.ToArray(), instance_name);
                            m_General.RowInfoList[row_idx].InstanceName = unique_name;
                            exist_names.Add(unique_name);
                        }
                        break;
                    default: break;
                }
            }


            /// <summary>
            /// Cast _val to _type by boxing.
            /// </summary>
            /// <param name="_type"></param>
            /// <param name="_val">string to be casted to _type.</param>
            /// <returns>
            /// System.Object boxing of _type
            /// If cast to _type is failed, returns null.
            /// </returns>
            private System.Object TypeCast(Type _type, string _val)
            {
                System.Object cast_result = null;

                if (_type == typeof(Char)) {
                    Char result;
                    if (Char.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(String)) {
                    cast_result = _val;
                } else if (_type == typeof(SByte)) {
                    SByte result;
                    if (SByte.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(Byte)) {
                    Byte result;
                    if (Byte.TryParse(_val, out result))
                        cast_result = _val;
                } else if (_type == typeof(Int16)) {
                    Int16 result;
                    if (Int16.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(UInt16)) {
                    UInt16 result;
                    if (UInt16.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(Int32)) {
                    Int32 result;
                    if (Int32.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(UInt32)) {
                    UInt32 result;
                    if (UInt32.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(Int64)) {
                    Int64 result;
                    if (Int64.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(UInt64)) {
                    UInt64 result;
                    if (UInt64.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(Single)) {
                    Single result;
                    if (Single.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(Double)) {
                    Double result;
                    if (!Double.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(Decimal)) {
                    Decimal result;
                    if (Decimal.TryParse(_val, out result))
                        cast_result = result;
                } else if (_type == typeof(Boolean)) {
                    Boolean result;
                    if (bool.TryParse(_val, out result))
                        cast_result = result;
                } else {
                    if (_type.IsEnum) {
                        if (Enum.IsDefined(_type, _val))
                            cast_result = Enum.Parse(_type, _val);
                    }
                }
                return cast_result;
            }

            /// <summary>
            /// Shows Instance List Part.
            /// </summary>
            /// <returns>
            /// - True: There is at least 1 instance which to be generated.
            /// - False: There is no instance which to be generated.
            /// </returns>
            private bool ShowInstanceListPart()
            {
                GUILayout.Space(5);
                GUILayout.Box("", GUILayout.ExpandWidth(true), GUILayout.Height(2));

                m_InstanceListPart.IsOpen = EditorGUILayout.Foldout(m_InstanceListPart.IsOpen,
                                                                    new GUIContent("Instance List"), m_General.FoldoutStyle);
                if (m_InstanceListPart.IsOpen) {
                    if (m_General.RowInfoList.Count == 0) {
                        EditorGUILayout.HelpBox("Not found any instance data which match with Script field.", MessageType.Warning);
                        return false;
                    }
                    else {
                        EditorGUILayout.BeginHorizontal();
                        {
                            EditorGUILayout.LabelField("Naming Rule");

                            GUILayout.FlexibleSpace();

                            EditorGUI.BeginChangeCheck();
                            m_InstanceListPart.NamingRule = (InstanceListPart.NamingRules)EditorGUILayout.EnumPopup(m_InstanceListPart.NamingRule, GUILayout.Width(150));
                            if (EditorGUI.EndChangeCheck())
                                SetInstanceName();
                        }
                        EditorGUILayout.EndHorizontal();

                        switch (m_InstanceListPart.NamingRule) {
                            case (InstanceListPart.NamingRules.BaseNameWithIndex):
                                EditorGUILayout.BeginHorizontal(); {

                                    EditorGUILayout.LabelField("Base Name");

                                    EditorGUI.BeginChangeCheck();
                                    // Adjust width depending on max length of instance name 
                                    float width = Mathf.Max(150, EditorStyles.label.CalcSize(new GUIContent(m_InstanceListPart.BaseName)).x + 10);
                                    m_InstanceListPart.BaseName = EditorGUILayout.TextField(m_InstanceListPart.BaseName, GUILayout.Width(width));
                                    if (EditorGUI.EndChangeCheck())
                                        SetInstanceName();
                                }
                                EditorGUILayout.EndHorizontal();
                                break;
                            case (InstanceListPart.NamingRules.FieldValue):
                                EditorGUILayout.BeginHorizontal(); {

                                    EditorGUILayout.LabelField("Reference Field");

                                    string[] field_names = m_General.FieldInfoArray.Select(x => x.Name).ToArray();
                                    EditorGUI.BeginChangeCheck();
                                    m_InstanceListPart.ReferenceFieldIndex = EditorGUILayout.Popup(m_InstanceListPart.ReferenceFieldIndex, field_names, GUILayout.Width(150));
                                    if (EditorGUI.EndChangeCheck())
                                        SetInstanceName();
                                }
                                EditorGUILayout.EndHorizontal();
                                break;
                            default: break;
                        }
                        ShowTable();
                    }
                }
                return true;
            }

            /// <summary>
            /// Shows Instance List Table.
            /// </summary>
            private void ShowTable()
            {
                int common_height = 25;

                // Access
                int access_width = 80;
                GUILayoutOption[] access_vertical_option = { GUILayout.Width(access_width) };
                GUILayoutOption[] access_item_option = { GUILayout.Width(access_width), GUILayout.Height(common_height) };

                // Data Type
                float type_width = 80;
                foreach (FieldInfo field in m_General.FieldInfoArray) {
                    // Adjust width depending on max length of field type
                    float req_width = EditorStyles.label.CalcSize(new GUIContent(ConvertToAliasTypeName(field.FieldType.ToString()))).x + 10;
                    if (type_width < req_width)
                        type_width = req_width;
                }
                GUILayoutOption[] type_vertical_option = { GUILayout.Width(type_width) };
                GUILayoutOption[] type_item_option = { GUILayout.Width(type_width), GUILayout.Height(common_height) };

                // Field Name
                float name_width = 100;
                foreach (FieldInfo field in m_General.FieldInfoArray) {
                    // Adjust width depending on max length of field name
                    float req_width = EditorStyles.label.CalcSize(new GUIContent(field.Name)).x + 10;
                    if (name_width < req_width)
                        name_width = req_width;
                }
                GUILayoutOption[] name_vertical_option = { GUILayout.Width(name_width), GUILayout.ExpandWidth(false) };
                GUILayoutOption[] name_item_option = { GUILayout.Width(name_width), GUILayout.ExpandWidth(false), GUILayout.Height(common_height) };

                // Instance value
                float value_width = 100;
                foreach (RowInfo row_info in m_General.RowInfoList) {
                    // Adjust width depending on max length of instance value
                    foreach (string val in row_info.Value) {
                        float req_width = EditorStyles.textField.CalcSize(new GUIContent(val)).x + 10;
                        if (value_width < req_width)
                            value_width = req_width;
                    }
                }
                GUILayoutOption[] value_vertical_option = { GUILayout.Width(value_width), GUILayout.ExpandWidth(true) };
                GUILayoutOption[] value_item_option = { GUILayout.Width(value_width), GUILayout.ExpandWidth(true), GUILayout.Height(common_height) };

                GUIStyle style = new GUIStyle("Box");
                style.normal.background = Texture2D.whiteTexture;

                EditorGUILayout.BeginVertical(GUI.skin.box, GUILayout.ExpandWidth(true));
                {
                    m_InstanceListPart.ScrollPosition = EditorGUILayout.BeginScrollView(m_InstanceListPart.ScrollPosition);

                    for (int row_idx = 1; row_idx < m_General.RowInfoList.Count; row_idx++) {
                        RowInfo current_row_info = m_General.RowInfoList[row_idx];
                        current_row_info.IsListItemOpen = EditorGUILayout.Foldout(current_row_info.IsListItemOpen,
                                                                                  new GUIContent(current_row_info.InstanceName),
                                                                                  m_General.FoldoutStyle);
                        if (current_row_info.IsListItemOpen) {
                            EditorGUILayout.BeginHorizontal(style, GUILayout.ExpandWidth(true));
                            {
                                // Access
                                EditorGUILayout.BeginVertical(access_vertical_option);
                                {
                                    foreach (FieldInfo field in m_General.FieldInfoArray) {
                                        EditorGUILayout.LabelField(field.Attributes.ToString().ToLower(), access_item_option);
                                    }
                                }
                                EditorGUILayout.EndVertical();

                                // Data Type 
                                EditorGUILayout.BeginVertical(type_vertical_option);
                                {
                                    foreach (FieldInfo field in m_General.FieldInfoArray) {
                                        EditorGUILayout.LabelField(ConvertToAliasTypeName(field.FieldType.ToString()), type_item_option);
                                    }
                                }
                                EditorGUILayout.EndVertical();

                                // Field Name
                                EditorGUILayout.BeginVertical(name_vertical_option);
                                {
                                    foreach (FieldInfo field in m_General.FieldInfoArray) {
                                        EditorGUILayout.LabelField(field.Name, name_item_option);
                                    }
                                }
                                EditorGUILayout.EndVertical();

                                // Instance Value
                                EditorGUILayout.BeginVertical(value_vertical_option);
                                {
                                    EditorGUI.BeginDisabledGroup(true);
                                    {
                                        foreach (FieldInfo field in m_General.FieldInfoArray) {
                                            int val_idx = m_General.RowInfoList[0].Value.IndexOf(field.Name);
                                            if (val_idx == -1) // Not found
                                                EditorGUILayout.TextField(string.Empty, value_item_option);
                                            else 
                                                EditorGUILayout.TextField(current_row_info.Value[val_idx], value_item_option);
                                        }
                                    }
                                    EditorGUI.EndDisabledGroup();
                                }
                                EditorGUILayout.EndVertical();
                            }
                            EditorGUILayout.EndHorizontal();
                        }
                    }
                    EditorGUILayout.EndScrollView();
                }
                EditorGUILayout.EndVertical();
            }

            /// <summary>
            /// Convert data type name from .Net to alias. (e.g., System.Int32 > int)
            /// </summary>
            /// <param name="_system_type_name">.Net data type name</param>
            /// <returns>
            /// Alias data type name.
            /// </returns>
            public string ConvertToAliasTypeName(string _system_type_name)
            {
                string alias = string.Empty;
                switch (_system_type_name) {
                    case "System.Boolean": alias = "bool"; break;
                    case "System.Byte":    alias = "byte"; break;
                    case "System.SByte":   alias = "sbyte"; break;
                    case "System.Char":    alias = "char"; break;
                    case "System.Decimal": alias = "decimal"; break;
                    case "System.Double":  alias = "double"; break;
                    case "System.Single":  alias = "float"; break;
                    case "System.Int32":   alias = "int"; break;
                    case "System.UInt32":  alias = "uint"; break;
                    case "System.Int64":   alias = "long"; break;
                    case "System.UInt64":  alias = "ulong"; break;
                    case "System.Object":  alias = "object"; break;
                    case "System.Int16":   alias = "short"; break;
                    case "System.UInt16":  alias = "ushort"; break;
                    case "System.String":  alias = "string"; break;
                    default: alias = _system_type_name; break;
                }
                return alias;
            }

            /// <summary>
            /// Shows Output Part.
            /// </summary>
            /// <returns>
            /// - True: Specified directory asset is valid.
            /// - False: Invalid directory is specified.
            /// </returns>
            private bool ShowOutputPart()
            {
                GUILayout.Space(5);
                GUILayout.Box("", GUILayout.ExpandWidth(true), GUILayout.Height(2));
                m_OutputPart.IsOpen = EditorGUILayout.Foldout(m_OutputPart.IsOpen,
                                                              new GUIContent("Output"), m_General.FoldoutStyle);
                if (m_OutputPart.IsOpen) {
                    // Directory Field
                    EditorGUILayout.BeginHorizontal();
                    {
                        EditorGUILayout.LabelField("Directory");
                        GUILayout.FlexibleSpace();
                        m_OutputPart.OutputDirAsset = EditorGUILayout.ObjectField(m_OutputPart.OutputDirAsset, typeof(UnityEngine.Object), false, GUILayout.Width(150));
                    }
                    EditorGUILayout.EndHorizontal();

                    string output_dir_path = AssetDatabase.GetAssetPath(m_OutputPart.OutputDirAsset);
                    if (AssetDatabase.IsValidFolder(output_dir_path))
                        m_OutputPart.OutputDirPath = output_dir_path;
                    else {
                        EditorGUILayout.HelpBox("Please specify a directory.", MessageType.Error);
                        return false;
                    }

                    // File name Label
                    EditorGUILayout.BeginHorizontal();
                    {
                        EditorGUILayout.LabelField("File Name");
                        GUILayout.FlexibleSpace();
                        EditorGUILayout.LabelField("As the above settings", GUILayout.Width(150));
                    }
                    EditorGUILayout.EndHorizontal();

                    // Generate Button
                    EditorGUILayout.BeginHorizontal();
                    {
                        GUILayout.FlexibleSpace();
                        if (GUILayout.Button("Generate", GUILayout.ExpandWidth(false), GUILayout.Width(100))) {
                            GenerateInstance();
                            if (m_ExcelPart.EVWindow != null)
                                m_ExcelPart.EVWindow.Close();
                        }
                        GUILayout.FlexibleSpace();
                    }
                    EditorGUILayout.EndHorizontal();
                }
                return true;
            }

            /// <summary>
            /// Generates ScruptableObject Instance.
            /// </summary>
            private void GenerateInstance()
            {
                List<string> existing_file_name = new List<string>();
                for (int row_idx = 1; row_idx < m_General.RowInfoList.Count; row_idx++) {
                    ScriptableObject ins = ScriptableObject.CreateInstance(m_ScriptPart.SOSAsset.name);
                    foreach (FieldInfo field in m_General.FieldInfoArray) {
                        int val_idx = m_General.RowInfoList[0].Value.IndexOf(field.Name);
                        if (val_idx != -1)
                            field.SetValue(ins, TypeCast(field.FieldType, m_General.RowInfoList[row_idx].Value[val_idx]));
                    }

                    string path = m_OutputPart.OutputDirPath;
                    AssetDatabase.CreateAsset(ins, path + "//" + m_General.RowInfoList[row_idx].InstanceName.TrimStart() + ".asset");
                    AssetDatabase.SaveAssets();
                    AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);
                }
                m_OutputPart.ReqReloadExcel = true; 
            }

            /// <summary>
            /// Shows Menu Button
            /// .</summary>
            /// <returns>
            /// - True: The button was not clicked.
            /// - False: The button was clicked.
            /// </returns>
            private bool ShowMenuButton()
            {
                GUILayout.Box("", GUILayout.ExpandWidth(true), GUILayout.Height(2));
                GUILayout.Space(15);

                EditorGUILayout.BeginHorizontal();
                {
                    GUILayout.FlexibleSpace();
                    if (GUILayout.Button("Return to Menu", GUILayout.ExpandWidth(false), GUILayout.Width(100))) {
                        if (m_ExcelPart.EVWindow != null)
                            m_ExcelPart.EVWindow.Close();

                        return true;
                    }
                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();
                GUILayout.Space(15);
                return false;
            }

            /// <summary>
            /// Callback function called when ExcelViewer set interested cells.
            /// </summary>
            /// <remarks>
            /// Updates RowInfoList according to the new range, and then call XL2SO's Repaint() to repaint FieldList.
            /// </remarks>
            /// <param name="_start_cell"> Start cell of selected cells</param>
            /// <param name="_last_cell"> Last cell of selected cells</param>
            private void OnSetSelectCells(Vector2Int _start_cell, Vector2Int _last_cell)
            {
                UpdateRowInfoList(_start_cell, _last_cell);
                m_Parent.Repaint();
            }
        }
    }
}
