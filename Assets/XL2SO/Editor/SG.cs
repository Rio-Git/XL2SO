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
using System.CodeDom.Compiler;

namespace XL2SO
{
    namespace SG
    {
        /// <summary>
        /// Represents a "Generate ScriptableObject Script" feature.
        /// </summary>
        public class SG : IScreen
        {
            private XL2SO         m_Parent        = null;
            private General       m_General       = null;
            private ExcelPart     m_ExcelPart     = null;
            private FieldListPart m_FieldListPart = null;
            private OutputPart    m_OutputPart    = null;

            /// <summary>
            /// Initializes a new instance of the <see cref="SG"/> class.
            /// </summary>
            /// <remarks>
            /// Set instance members and check if template file exists.
            /// </remarks>
            /// <param name="_Parent"> Parent <see cref="XL2SO"/> instance </param>
            override public void Initialize(XL2SO _Parent)
            {
                // Initialize instance members
                m_Parent        = _Parent;
                m_General       = new General();
                m_ExcelPart     = new ExcelPart();
                m_FieldListPart = new FieldListPart();
                m_OutputPart    = new OutputPart();
                CreateDataTypeList();
                CreateGUIStyles();

                // Check if template file exists
                MonoScript mono_this = MonoScript.FromScriptableObject(m_Parent);
                string script_path = AssetDatabase.GetAssetPath(mono_this);
                string script_directory = Path.GetDirectoryName(script_path);
                m_General.TemplatePath = script_directory + "\\ScriptTemplate.txt";
                m_General.TemplateExists = File.Exists(m_General.TemplatePath);
            }

            /// <summary>
            /// Creates data type list from PrimaryDataTypes and User-defined enum.
            /// </summary>
            private void CreateDataTypeList()
            {
                m_FieldListPart.DataTypeList = new List<string>();

                // Add PrimaryDataTypes with type prefix
                foreach (string enum_name in Enum.GetNames(typeof(PrimaryDataTypes))) {
                    // Convert string to enum for avoiding typo error
                    PrimaryDataTypes converted;
                    if (Enum.TryParse(enum_name, true, out converted)) {
                        switch (converted) {
                            case (PrimaryDataTypes.@sbyte):
                            case (PrimaryDataTypes.@byte):
                            case (PrimaryDataTypes.@short):
                            case (PrimaryDataTypes.@ushort):
                            case (PrimaryDataTypes.@int):
                            case (PrimaryDataTypes.@uint):
                            case (PrimaryDataTypes.@long):
                            case (PrimaryDataTypes.@ulong):
                                m_FieldListPart.DataTypeList.Add(General.IntPrefix + enum_name);
                                break;
                            case (PrimaryDataTypes.@char):
                            case (PrimaryDataTypes.@string):
                                m_FieldListPart.DataTypeList.Add(General.CharPrefix + enum_name);
                                break;
                            case (PrimaryDataTypes.@float):
                            case (PrimaryDataTypes.@double):
                            case (PrimaryDataTypes.@decimal):
                                m_FieldListPart.DataTypeList.Add(General.RealPrefix + enum_name);
                                break;
                            default:
                                m_FieldListPart.DataTypeList.Add(enum_name);
                                break;
                        }
                    }
                }

                // Add User-defined Enum
                List<string> user_defined_enum = new List<string>();
                Assembly[] current_assemblies = AppDomain.CurrentDomain.GetAssemblies();
                Assembly target_assembly = current_assemblies.SingleOrDefault(x => x.GetName().Name == "Assembly-CSharp");
                if (target_assembly != null) {
                    foreach (Type type in target_assembly.GetExportedTypes()) {
                        if (type.IsEnum) {
                            // Type.FullName contains its name like "[Class Name]+[Enum Name]" 
                            // for inner enum which is defined in a class
                            string plus2dot_text = type.FullName.Replace("+", ".");
                            m_FieldListPart.DataTypeList.Add(General.EnumPrefix + plus2dot_text);
                        }
                    }
                }
            }

            /// <summary>
            /// Creates GUIStyles for some GUI components.
            /// </summary>
            private void CreateGUIStyles()
            {
                m_General.FoldoutStyle = new GUIStyle("FoldOut");
                m_General.FoldoutStyle.fontStyle = FontStyle.Bold;

                m_FieldListPart.ColumnStyle = new GUIStyle("Box");
                Texture2D dot_bg = new Texture2D(1, 1);
                dot_bg.SetPixel(1, 1, new Color(0.25f, 0.25f, 0.25f));
                dot_bg.Apply();
                m_FieldListPart.ColumnStyle.normal.background = dot_bg;
                m_FieldListPart.ColumnStyle.normal.textColor = Color.white;
                m_FieldListPart.ColumnStyle.margin = new RectOffset(0, 0, 0, 0);

                m_FieldListPart.ContentStyle = new GUIStyle("Box");
                m_FieldListPart.ContentStyle.normal.background = Texture2D.whiteTexture;
                m_FieldListPart.ContentStyle.margin = new RectOffset(0, 0, 0, 0);
                m_FieldListPart.PopupStyle = new GUIStyle(EditorStyles.popup);
                m_FieldListPart.PopupStyle.alignment = TextAnchor.MiddleCenter;
            }

            /// <summary>
            /// Main routine of <see cref="SG"/> package.
            /// </summary>
            /// <returns>
            /// A new instance of <see cref="Menu"/> class: When "Return to Menu" button is clicked.
            /// A instance of <see cref="SG"/> class (this): Other than above case.
            /// </returns>
            override public IScreen Display()
            {
                if (m_General.TemplateExists) {
                    m_General.ScrollPosition = EditorGUILayout.BeginScrollView(m_General.ScrollPosition);
                    if (ShowExcelPart()) {
                        if (ShowFieldListPart())
                            ShowOutputPart();
                    }
                    EditorGUILayout.EndScrollView();
                } else 
                    EditorGUILayout.HelpBox("Not found template file", MessageType.Error);

                if (ShowMenuButton())
                    return CreateInstance<Menu>();
                else
                    return this;
            }

            /// <summary>
            /// Shows Excel Part.
            /// </summary>
            /// <remarks>
            /// Error checkers:
            /// - Format check: The specified Excel asset is .xls or .xlsx file or not.
            /// - Contents check: The Excel has at least 1 sheet which contains at least 1 row (for column label) or not.
            /// 
            /// Flag Combination Table: 
            /// - IsStatup | IsValidExcel | IsExcelAsset   
            /// -     1    |       0      |       0      => Excel asset has never changed.
            /// -     0    |       0      |       0      => Format error.
            /// -     0    |       0      |       1      => Contents error.
            /// -     0    |       1      |       1      => Complete all process correctly.
            /// 
            /// Request for reloading Excel:
            /// - When generating script is done by GenerateScript(), Unity do serialization/deserializatoin unexpectedly and 
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
                            for (int sheet_idx = 0; sheet_idx < m_ExcelPart.Book.NumberOfSheets; sheet_idx++) {
                                ISheet sheet = m_ExcelPart.Book.GetSheetAt(sheet_idx);
                                if (sheet.PhysicalNumberOfRows > 0) {
                                    m_ExcelPart.SheetNameList.Add(sheet.SheetName);
                                }
                            }

                            if (m_ExcelPart.SheetNameList.Count > 0) {
                                ISheet select_sheet = m_ExcelPart.Book.GetSheet(m_ExcelPart.SheetNameList[m_ExcelPart.SelectedSheetIndex]);
                                m_ExcelPart.Sheet = select_sheet;
                                InitializeColumnInfoList();

                                // Set default values used in OutputPart
                                m_OutputPart.OutputFileNameWithoutExtension = m_ExcelPart.Sheet.SheetName;
                                string excel_dir_path = Path.GetDirectoryName(m_ExcelPart.ExcelPath);
                                UnityEngine.Object excel_dir_asset = AssetDatabase.LoadAssetAtPath(excel_dir_path, typeof(UnityEngine.Object));
                                m_OutputPart.OutputDirAsset = excel_dir_asset;
                                m_OutputPart.OutputDirPath = excel_dir_path;

                                m_ExcelPart.IsValidExcel = true;
                            } else
                                m_ExcelPart.IsValidExcel = false;
                        } else {
                            m_ExcelPart.IsExcelAsset = false;
                            m_ExcelPart.IsValidExcel = false;
                        }
                    }

                    // Display Error Message
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

                                m_ExcelPart.Sheet =  m_ExcelPart.Book.GetSheet(m_ExcelPart.SheetNameList[m_ExcelPart.SelectedSheetIndex]);
                                InitializeColumnInfoList();
                                
                                m_OutputPart.OutputFileNameWithoutExtension = m_ExcelPart.Sheet.SheetName;
                            }
                        }
                        EditorGUILayout.EndHorizontal();

                        // Show Cells Label
                        EditorGUILayout.BeginHorizontal();
                        {
                            EditorGUILayout.LabelField("Cells");
                            string format = "({0},{1}) to ({2},{3})";
                            int start_x   = Mathf.Min(m_ExcelPart.StartCell.x, m_ExcelPart.LastCell.x);
                            int start_y   = Mathf.Min(m_ExcelPart.StartCell.y, m_ExcelPart.LastCell.y);
                            int end_x     = Mathf.Max(m_ExcelPart.StartCell.x, m_ExcelPart.LastCell.x);
                            int end_y     = Mathf.Max(m_ExcelPart.StartCell.y, m_ExcelPart.LastCell.y);
                            string cells  = string.Format(format, start_x, start_y, end_x, end_y);
                            EditorGUILayout.LabelField(cells, GUILayout.Width(150));
                        }
                        EditorGUILayout.EndHorizontal();

                        // Show Cells Label
                        EditorGUILayout.BeginHorizontal();
                        {
                            GUILayout.FlexibleSpace();
                            if (GUILayout.Button("Select Cells", GUILayout.Width(150))) {

                                // Reload Excel file after the script is generated 
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
            /// Initializes ColumnInfoList with a default range.
            /// </summary>
            /// <remarks>
            /// Default range is set as below:
            /// - Column: FirstRow.FirstCellNum - FirstRow.LastCellNum
            /// - Row: FirstRowNum - LastRowNum
            /// This means ColumnInfoList includes all cells in the sheet 
            /// </remarks>
            private void InitializeColumnInfoList()
            {
                IRow first_row    = m_ExcelPart.Sheet.GetRow(m_ExcelPart.Sheet.FirstRowNum);
                int start_col_idx = first_row.FirstCellNum;
                int start_row_idx = m_ExcelPart.Sheet.FirstRowNum;
                int last_col_idx  = first_row.LastCellNum - 1; // LastCellNum equals to total num of cells
                int last_row_idx  = m_ExcelPart.Sheet.LastRowNum;
                UpdateColumnInfoList(new Vector2Int(start_col_idx, start_row_idx), new Vector2Int(last_col_idx, last_row_idx));
            }

            /// <summary>
            /// Updates ColumnInfoList by a range of _start_cell and _start_cell.
            /// </summary>
            /// <param name="_start_cell"> start cell of selected cells </param>
            /// <param name="_last_cell"> last cell of selected cells</param>
            public void UpdateColumnInfoList(Vector2Int _start_cell, Vector2Int _last_cell)
            {
                int start_col_idx = Mathf.Min(_start_cell.x, _last_cell.x);
                int last_col_idx  = Mathf.Max(_start_cell.x, _last_cell.x);
                int start_row_idx = Mathf.Min(_start_cell.y, _last_cell.y);
                int last_row_idx  = Mathf.Max(_start_cell.y, _last_cell.y);

                m_General.ColumnInfoList = new List<ColumnInfo>();
                //for (int col_idx = start_col_idx; col_idx <= last_col_idx; col_idx++) {
                for (int col_idx = start_col_idx; col_idx <= last_col_idx; col_idx++) {
                    ColumnInfo new_col_info = new ColumnInfo();
                    for (int row_idx = start_row_idx; row_idx <= last_row_idx; row_idx++) {
                        IRow row = m_ExcelPart.Sheet.GetRow(row_idx);
                        if (row == null) // Blank row
                            new_col_info.Value.Add(string.Empty);
                        else {
                            ICell cell = row.GetCell(col_idx);
                            if (cell != null)
                                new_col_info.Value.Add(cell.ToString());
                            else // Blank cell
                                new_col_info.Value.Add(string.Empty);
                        }
                    }
                    m_General.ColumnInfoList.Add(new_col_info);
                }
                m_ExcelPart.StartCell = _start_cell;
                m_ExcelPart.LastCell  = _last_cell;
                SetDefaultProperty();
            }

            /// <summary>
            /// Sets Default properties (Ignore, Type and Access) for each ColumnnInfo in CoulmInfoList.
            /// </summary>
            /// <remarks>
            /// The Default properties depends on its column label (ColumnInfo.Value[0]) and data as below:
            /// - The label string is incorrect as C# format
            ///   > Ignore = true, Access = public, Type = string
            /// - ColumnInfo.Value.Count = 1
            ///   > Ignore = false, Access = public, Type = string
            /// - If all data in ColumnInfo.Value other than [0] can be parsed to Int
            ///   > Ignore = false, Access = public, Type = int
            /// - Other than above cases
            ///   > Ignore = false, Access = public, Type = string
            /// </remarks>
            private void SetDefaultProperty()
            {
                for (int col_idx = 0; col_idx < m_General.ColumnInfoList.Count; col_idx++) {
                    ColumnInfo col_info = m_General.ColumnInfoList[col_idx];
                    CodeDomProvider CDprovider = CodeDomProvider.CreateProvider("C#");

                    if (CDprovider.IsValidIdentifier(col_info.Value[0])) {
                        col_info.Access = AccessLevels.@public;
                        col_info.Ignore = false;
                        if (col_info.Value.Count == 1)
                            col_info.Type = General.CharPrefix + "string";
                        else {
                            int row_num = Mathf.Abs(m_ExcelPart.StartCell.y - m_ExcelPart.LastCell.y);
                            for (int row_idx = 1; row_idx < row_num; row_idx++) {
                                if (Int32.TryParse(col_info.Value[row_idx], out int temp))
                                    col_info.Type = General.IntPrefix + "int";
                                else {
                                    col_info.Type = General.CharPrefix + "string";
                                    break;
                                }
                            }
                        }
                    } else {
                        col_info.Ignore = true;
                        col_info.Access = AccessLevels.@public;
                        col_info.Type = General.CharPrefix + "string";
                    }
                }

                if (!m_General.ColumnInfoList.Any(x => x.Ignore == false))
                    m_FieldListPart.IgnoreAll = true;

                m_FieldListPart.TypeAll = General.CharPrefix + "string";
                m_FieldListPart.AccessAll = AccessLevels.@public;
            }


            /// <summary>
            /// Shows Field List Part.
            /// </summary>
            /// <returns>
            /// - True: There is at least 1 field which to be generated.
            /// - False: There is no field which to be generated.
            /// </returns>
            private bool ShowFieldListPart()
            {
                GUILayout.Space(5);
                GUILayout.Box("", GUILayout.ExpandWidth(true), GUILayout.Height(2));
                m_FieldListPart.IsOpen = EditorGUILayout.Foldout(m_FieldListPart.IsOpen, new GUIContent("Field List"), m_General.FoldoutStyle);
                GUILayout.Space(5);
                if (m_FieldListPart.IsOpen)
                    ShowTable();

               if (m_FieldListPart.IgnoreAll == true) {
                   EditorGUILayout.HelpBox("Please set at least 1 field as Ignore = false.", MessageType.Error);
                   return false;
               } else
                    return true;
            }

            /// <summary>
            /// Shows Field List Table.
            /// </summary>
            private void ShowTable()
            {
                int common_height = 25;

                // Ignore
                int ignore_width = 60;
                GUILayoutOption[] ignore_vertical_option = { GUILayout.Width(ignore_width) };
                GUILayoutOption[] ignore_item_option     = { GUILayout.Width(ignore_width), GUILayout.Height(common_height) };

                // Access
                int access_width = 80;
                GUILayoutOption[] access_vertical_option = { GUILayout.Width(access_width) };
                GUILayoutOption[] access_item_option     = { GUILayout.Width(access_width), GUILayout.Height(common_height) };

                // Data Type
                float type_width = 80;
                // Adjust width depending on max length of ColumnInfo.Type in ColumnInfoList
                foreach (ColumnInfo col_info in m_General.ColumnInfoList) {
                    float req_width = EditorStyles.popup.CalcSize(new GUIContent(TrimTypePrefix(col_info.Type))).x + 10;
                    if (type_width < req_width)
                        type_width = req_width;
                }
                GUILayoutOption[] type_vertical_option = { GUILayout.Width(type_width) };
                GUILayoutOption[] type_item_option     = { GUILayout.Width(type_width), GUILayout.Height(common_height) };

                // Field Name
                float name_width = 70;
                // Adjust width depending on max length of ColumnInfo.Value in ColumnInfoList
                foreach (ColumnInfo col_info in m_General.ColumnInfoList) {
                    float req_width = EditorStyles.popup.CalcSize(new GUIContent(TrimTypePrefix(col_info.Value[0]))).x + 10;
                    if (name_width < req_width)
                        name_width = req_width;
                }
                GUILayoutOption[] name_vertical_option = { GUILayout.Width(name_width), GUILayout.ExpandWidth(true) };
                GUILayoutOption[] name_item_option     = { GUILayout.Width(name_width), GUILayout.ExpandWidth(true),
                                                           GUILayout.Height(common_height) };

                EditorGUILayout.BeginHorizontal();
                {
                    GUILayout.Space(10);
                    EditorGUILayout.BeginVertical();
                    {
                        // Column
                        EditorGUILayout.BeginHorizontal(m_FieldListPart.ColumnStyle, GUILayout.ExpandWidth(true));
                        {
                            // Ignore
                            EditorGUILayout.BeginVertical(ignore_vertical_option);
                            {
                                EditorGUILayout.LabelField("Ignore", EditorStyles.whiteLabel, ignore_item_option);
                                GUILayout.Space(-4);
                                EditorGUI.BeginChangeCheck();
                                EditorGUILayout.BeginHorizontal();
                                {
                                    GUILayout.Space(20);
                                    m_FieldListPart.IgnoreAll = EditorGUILayout.Toggle(m_FieldListPart.IgnoreAll, ignore_item_option);
                                    GUILayout.Space(-20);
                                }
                                EditorGUILayout.EndHorizontal();
                                if (EditorGUI.EndChangeCheck())
                                    m_General.ColumnInfoList.ForEach(x => x.Ignore = m_FieldListPart.IgnoreAll);

                            }
                            EditorGUILayout.EndVertical();

                            GUILayout.Space(10);

                            // Access 
                            EditorGUILayout.BeginVertical(access_vertical_option);
                            {
                                EditorGUILayout.LabelField("Accessibility", EditorStyles.whiteLabel, access_item_option);
                                EditorGUI.BeginDisabledGroup(m_FieldListPart.IgnoreAll);
                                {
                                    EditorGUI.BeginChangeCheck();
                                    m_FieldListPart.AccessAll = (AccessLevels)EditorGUILayout.EnumPopup(m_FieldListPart.AccessAll, m_FieldListPart.PopupStyle, access_item_option);
                                    if (EditorGUI.EndChangeCheck())
                                        m_General.ColumnInfoList.ForEach(x => x.Access = m_FieldListPart.AccessAll);
                                }
                                EditorGUI.EndDisabledGroup();
                            }
                            EditorGUILayout.EndVertical();

                            GUILayout.Space(20);

                            // Data Type 
                            EditorGUILayout.BeginVertical(type_vertical_option);
                            {
                                EditorGUILayout.BeginHorizontal();
                                {
                                    // Centering label 
                                    GUILayout.Space((type_width - 60) / 2);
                                    EditorGUILayout.LabelField("Data Type", EditorStyles.whiteLabel, type_item_option);
                                    GUILayout.Space(-(type_width - 60) / 2);
                                }
                                EditorGUILayout.EndHorizontal();

                                EditorGUI.BeginDisabledGroup(m_FieldListPart.IgnoreAll);
                                {
                                    string shrink_type = TrimTypePrefix(m_FieldListPart.TypeAll);
                                    if (GUILayout.Button(shrink_type, m_FieldListPart.PopupStyle, type_item_option)) {
                                        GenericMenu type_menu = new GenericMenu();
                                        foreach (string item in m_FieldListPart.DataTypeList) {
                                            type_menu.AddItem(new GUIContent(item), item.Equals(m_FieldListPart.TypeAll),
                                            (() =>
                                               {
                                                   m_FieldListPart.TypeAll = item;
                                                   m_General.ColumnInfoList.ForEach(x => x.Type = m_FieldListPart.TypeAll);
                                               }
                                            ));
                                        }
                                        type_menu.ShowAsContext();
                                    }
                                }
                                EditorGUI.EndDisabledGroup();
                            }
                            EditorGUILayout.EndVertical();

                            GUILayout.Space(20);

                            // Field Name
                            EditorGUILayout.BeginVertical(name_vertical_option);
                            {
                                EditorGUILayout.LabelField("Name", EditorStyles.whiteLabel, name_item_option);
                                EditorGUILayout.LabelField("", name_item_option);
                            }
                            EditorGUILayout.EndVertical();
                        }
                        EditorGUILayout.EndHorizontal();

                        GUILayout.Space(-10);

                        // Contents
                        m_FieldListPart.ScrollPosition = EditorGUILayout.BeginScrollView(m_FieldListPart.ScrollPosition);
                        EditorGUILayout.BeginHorizontal(m_FieldListPart.ContentStyle, GUILayout.ExpandWidth(true));
                        {
                            EditorGUILayout.BeginVertical();
                            {
                                GUILayout.Space(8);
                                EditorGUILayout.BeginHorizontal();
                                {
                                    // Ignore
                                    EditorGUILayout.BeginVertical(ignore_vertical_option);
                                    {
                                        foreach (ColumnInfo col_info in m_General.ColumnInfoList) {
                                            GUILayout.Space(-4);
                                            EditorGUILayout.BeginHorizontal();
                                            {
                                                GUILayout.Space(20);
                                                col_info.Ignore = EditorGUILayout.Toggle(col_info.Ignore, ignore_item_option);
                                                GUILayout.Space(-20);
                                            }
                                            EditorGUILayout.EndHorizontal();
                                            GUILayout.Space(4);
                                        }
                                    }
                                    EditorGUILayout.EndVertical();

                                    GUILayout.Space(14);

                                    // Access 
                                    EditorGUILayout.BeginVertical(access_vertical_option);
                                    {
                                        foreach (ColumnInfo col_info in m_General.ColumnInfoList) {
                                            EditorGUI.BeginDisabledGroup(col_info.Ignore);
                                            {
                                                col_info.Access = (AccessLevels)EditorGUILayout.EnumPopup(col_info.Access,
                                                                                                          m_FieldListPart.PopupStyle, access_item_option);
                                            }
                                            EditorGUI.EndDisabledGroup();
                                        }
                                    }
                                    EditorGUILayout.EndVertical();

                                    GUILayout.Space(25);

                                    // Data Type 
                                    EditorGUILayout.BeginVertical(type_vertical_option);
                                    {
                                        foreach (ColumnInfo col_info in m_General.ColumnInfoList) {
                                            string shrink_type = TrimTypePrefix(col_info.Type);
                                            EditorGUI.BeginDisabledGroup(col_info.Ignore);
                                            {
                                                if (GUILayout.Button(shrink_type, m_FieldListPart.PopupStyle, type_item_option)) {
                                                    GenericMenu type_menu = new GenericMenu();
                                                    foreach (string type_name in m_FieldListPart.DataTypeList) {
                                                        type_menu.AddItem(new GUIContent(type_name), 
                                                                          type_name.Equals(col_info.Type),
                                                                          (() => { col_info.Type = type_name; }));
                                                    }
                                                    type_menu.ShowAsContext();
                                                }
                                            }
                                            EditorGUI.EndDisabledGroup();
                                        }
                                    }
                                    EditorGUILayout.EndVertical();

                                    GUILayout.Space(20);

                                    // Field Name
                                    EditorGUILayout.BeginVertical(name_vertical_option);
                                    {
                                        foreach (ColumnInfo col_info in m_General.ColumnInfoList) {
                                            GUILayout.Space(-3);
                                            EditorGUI.BeginDisabledGroup(col_info.Ignore);
                                            {
                                                EditorGUILayout.LabelField(col_info.Value[0], name_item_option);
                                            }
                                            EditorGUI.EndDisabledGroup();
                                            GUILayout.Space(3);
                                        }
                                    }
                                    EditorGUILayout.EndVertical();
                                }
                                EditorGUILayout.EndHorizontal();
                            }
                            EditorGUILayout.EndVertical();
                        }
                        EditorGUILayout.EndHorizontal();
                        EditorGUILayout.EndScrollView();

                    }
                    EditorGUILayout.EndVertical();
                    GUILayout.Space(10);
                }
                EditorGUILayout.EndHorizontal();

                // If all CoulmnInfo in ColumnInfoList is marked as ignore, set IgnoreAll = true
                bool false_exists = m_General.ColumnInfoList.Any(x => x.Ignore == false);
                m_FieldListPart.IgnoreAll = !false_exists;
            }

            /// <summary>
            /// Removes type prefix from _text.
            /// </summary>
            /// <param name="_text"> The text to be trimmed.</param>
            /// <returns>The string that remains after type prefix is removed.</returns>
            private string TrimTypePrefix(string _text)
            {
                string shrink_type = string.Empty;
                if (_text.Contains(General.CharPrefix))
                    shrink_type = _text.Replace(General.CharPrefix, "");
                else if (_text.Contains(General.IntPrefix))
                    shrink_type = _text.Replace(General.IntPrefix, "");
                else if (_text.Contains(General.RealPrefix))
                    shrink_type = _text.Replace(General.RealPrefix, "");
                else if (_text.Contains(General.EnumPrefix))
                    shrink_type = _text.Replace(General.EnumPrefix, "");
                else
                    shrink_type = _text;

                return shrink_type;
            }

            /// <summary>
            /// Shows Output Part.
            /// </summary>
            /// <returns>
            /// - True: Asset format is correct.
            /// - False: Asset format is incorrect.
            /// </returns>
            private bool ShowOutputPart()
            {
                GUILayout.Space(5);
                GUILayout.Box("", GUILayout.ExpandWidth(true), GUILayout.Height(2));

                m_OutputPart.IsOpen = EditorGUILayout.Foldout(m_OutputPart.IsOpen,
                                                              new GUIContent("Output"), m_General.FoldoutStyle);
                if (m_OutputPart.IsOpen) {
                    EditorGUI.indentLevel++;
                    
                    // Directory Field 
                    EditorGUILayout.BeginHorizontal();
                    {
                        EditorGUILayout.LabelField("Directory");
                        GUILayout.FlexibleSpace();
                        m_OutputPart.OutputDirAsset = EditorGUILayout.ObjectField(m_OutputPart.OutputDirAsset,
                                                                                  typeof(UnityEngine.Object),
                                                                                  false,
                                                                                  GUILayout.Width(195));
                    }
                    EditorGUILayout.EndHorizontal();

                    // Check asset format
                    string output_dir_path = AssetDatabase.GetAssetPath(m_OutputPart.OutputDirAsset);
                    if (AssetDatabase.IsValidFolder(output_dir_path))
                        m_OutputPart.OutputDirPath = output_dir_path;
                    else {
                        EditorGUILayout.HelpBox("Please specify a directory.", MessageType.Error);
                        return false;
                    }

                    // File Name Field
                    EditorGUILayout.BeginHorizontal();
                    {
                        EditorGUILayout.LabelField("File Name");
                        GUILayout.FlexibleSpace();
                        EditorGUILayout.BeginHorizontal();
                        {
                            m_OutputPart.OutputFileNameWithoutExtension = 
                                EditorGUILayout.TextField(m_OutputPart.OutputFileNameWithoutExtension,
                                                          GUILayout.Width(165),
                                                          GUILayout.ExpandWidth(true));
                            GUILayout.Space(-15);
                            EditorGUILayout.LabelField(".cs", GUILayout.Width(40));
                        }
                        EditorGUILayout.EndHorizontal();
                    }
                    EditorGUILayout.EndHorizontal();
                    EditorGUI.indentLevel--;

                    // Generate Button
                    EditorGUILayout.BeginHorizontal();
                    {
                        GUILayout.FlexibleSpace();
                        EditorGUI.BeginDisabledGroup(m_FieldListPart.IgnoreAll);
                        {
                            if (GUILayout.Button("Generate", GUILayout.ExpandWidth(false), GUILayout.Width(100))) {
                                GenerateScript();
                                if (m_ExcelPart.EVWindow != null)
                                    m_ExcelPart.EVWindow.Close();
                            }
                        }
                        EditorGUI.EndDisabledGroup();
                        GUILayout.FlexibleSpace();
                    }
                    EditorGUILayout.EndHorizontal();
                }
                return true;
            }

            /// <summary>
            /// Generates ScruptableObject Script.
            /// </summary>
            private void GenerateScript()
            {
                List<string> new_lines = new List<string>();
                foreach (ColumnInfo col_info in m_General.ColumnInfoList) {
                    if (col_info.Ignore == false) {
                        string format = null;
                        if (col_info != m_General.ColumnInfoList.Last())
                            format = "    {0} {1} {2};" + System.Environment.NewLine;
                        else
                            format = "    {0} {1} {2};";

                        new_lines.Add(string.Format(format, col_info.Access.ToString(), TrimTypePrefix(col_info.Type), col_info.Value[0]));
                    }
                }

                string template_text = File.ReadAllText(m_General.TemplatePath);
                string replaced_text;
                replaced_text = template_text.Replace("\r\n", "\n").Replace("\n", System.Environment.NewLine);
                replaced_text = replaced_text.Replace("$ClassName$", m_OutputPart.OutputFileNameWithoutExtension);
                replaced_text = replaced_text.Replace("$Contents$", String.Join("", new_lines));
                string full_path = m_OutputPart.OutputDirPath + "//" + m_OutputPart.OutputFileNameWithoutExtension + ".cs";

                File.WriteAllText(full_path, replaced_text);
                AssetDatabase.ImportAsset(full_path);
                AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);
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

            private void OnDestroy()
            {
                if (m_ExcelPart.EVWindow)
                    m_ExcelPart.EVWindow.Close();
            }

            /// <summary>
            /// Callback function called when ExcelViewer set interested cells.
            /// </summary>
            /// <remarks>
            /// Updates ColumnInfoList according to the new range, and then call XL2SO's Repaint() to repaint FieldList.
            /// </remarks>
            /// <param name="_start_cell"> Start cell of selected cells</param>
            /// <param name="_last_cell"> Last cell of selected cells</param>
            private void OnSetSelectCells(Vector2Int _start_cell, Vector2Int _last_cell)
            {
                UpdateColumnInfoList(_start_cell, _last_cell);
                m_Parent.Repaint();
            }
        }
    }
}