using UnityEngine;
using UnityEditor;
using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace XL2SO
{
    /// <summary>
    /// Manage Excel viewer window.
    /// </summary>
    public class ExcelViewer : EditorWindow
    {
        private Vector2        m_ScrollPosition     = Vector2.zero;     // [GUI] Current scroll position
        private List<string>   m_CellTable          = null;             // [GUI] Cell table of given sheet
        private int            m_CellWidth          = 100;              // [GUI] Cell width
        private int            m_CellHeight         = 25;               // [GUI] Cell Height
        private int            m_RowNumWithLabel    = 0;                // [GUI] Total row num including column label
        private int            m_ColumnNumWithLabel = 0;                // [GUI] Total column num including row label
        private int            m_SelectIndex        = 0;                // [GUI] Select cell index in SelectionGrid
        private Vector2Int     m_StartCell          = Vector2Int.zero;  // [GUI] Start cell position
        private Rect           m_StartCellRect      = new Rect();       // [GUI] Rect of Start cell
        private Vector2Int     m_LastCell           = Vector2Int.zero;  // [GUI] Last cell position
        private bool           m_ShowFrame          = true;             // Show frame or not
        private GUIStyleHolder m_GUIStyleHolder     = null;             // Instance of GUIStyleHolder

        private Action<Vector2Int, Vector2Int> m_SetSelectCellsCB = null; // Callback called when selecting cells.

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelViewer"/> class.
        /// </summary>
        /// <param name="_start_cell">Initial start cell.</param>
        /// <param name="_last_cell">Initial last cell.</param>
        /// <param name="_source_sheet">Source sheet which contain cell data.</param>
        /// <param name="_cb">Callback function called when interested cells are selected. </param>
        public void Initialize(Vector2Int _start_cell,
                               Vector2Int _last_cell,
                               ISheet _source_sheet,
                               Action<Vector2Int, Vector2Int> _cb)
        {
            // LactCellNum equals to total num of columns 
            int last_col_idx_in_sheet = 0;
            for (int row_idx = _source_sheet.FirstRowNum; row_idx <= _source_sheet.LastRowNum; row_idx++) {
                IRow row = _source_sheet.GetRow(row_idx);
                if (last_col_idx_in_sheet < row.LastCellNum)
                    last_col_idx_in_sheet = row.LastCellNum;
            }
            m_ColumnNumWithLabel = last_col_idx_in_sheet + 1;
            
            // LastRowNum + 1 equals to total num of columns 
            m_RowNumWithLabel = (_source_sheet.LastRowNum + 1) + 1;

            // Adjust _start_cell and _last_cell with column/row label.
            m_SelectIndex = (_start_cell.x + 1) + (_start_cell.y + 1) * m_ColumnNumWithLabel;
            m_StartCell = new Vector2Int(_start_cell.x + 1, _start_cell.y + 1);
            int start_rect_x =  m_StartCell.x * m_CellWidth;
            int start_rect_y =  m_StartCell.y * m_CellHeight;
            m_StartCellRect = new Rect(start_rect_x, start_rect_y, m_CellWidth, m_CellHeight);

            m_LastCell = new Vector2Int(_last_cell.x + 1, _last_cell.y + 1);

            m_CellTable = new List<string>();
            for (int row_idx = -1; row_idx < m_RowNumWithLabel - 1; row_idx++) {
                IRow row = null;
                if (row_idx != -1) // Not column label
                    row = _source_sheet.GetRow(row_idx);

                for (int col_idx = -1; col_idx < m_ColumnNumWithLabel - 1; col_idx++) {
                    if (row_idx == -1) // column label
                        m_CellTable.Add(string.Empty);
                    else {
                        if (row == null) // Blank row
                            m_CellTable.Add(string.Empty);
                        else if (col_idx == -1) // row label
                            m_CellTable.Add(string.Empty);
                        else {
                            ICell cell = row.GetCell(col_idx);
                            if (cell != null)
                                m_CellTable.Add(cell.ToString());
                            else
                                m_CellTable.Add(string.Empty);
                        }
                    }
                }
            }

            m_GUIStyleHolder = CreateInstance<GUIStyleHolder>();
            m_GUIStyleHolder.Initialize();

            m_SetSelectCellsCB = _cb;
        }

        /// <summary>
        /// OnGUI() function of Monobehaviour.
        /// </summary>
        void OnGUI()
        {
            m_ScrollPosition = EditorGUILayout.BeginScrollView(m_ScrollPosition);

            DrawCellTable();
            ColorLabels();
            ColorCell();

            EditorGUILayout.EndScrollView();

            EditorGUILayout.BeginVertical();
            {
                GUILayout.Space(10);

                // Buttons
                EditorGUILayout.BeginHorizontal();
                {
                    EditorGUILayout.LabelField("Left Click:", EditorStyles.boldLabel, GUILayout.Width(70));
                    EditorGUILayout.LabelField("Set start cell of a range");
                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();

                EditorGUILayout.BeginHorizontal();
                {
                    EditorGUILayout.LabelField("Right Click:", EditorStyles.boldLabel, GUILayout.Width(80));
                    EditorGUILayout.LabelField("Set last cell of a range");
                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();

                // Buttons
                EditorGUILayout.BeginHorizontal();
                {
                    EditorGUI.BeginDisabledGroup(!m_ShowFrame);
                    {
                        GUILayout.FlexibleSpace();
                        if (GUILayout.Button("Set", GUILayout.ExpandWidth(false), GUILayout.Width(100))) {
                            // Adjust cell index to non-row/column label
                            m_SetSelectCellsCB?.Invoke(m_StartCell - new Vector2Int(1, 1), m_LastCell - new Vector2Int(1, 1));
                        }
                    }
                    EditorGUI.EndDisabledGroup();

                    GUILayout.FlexibleSpace();

                    if (GUILayout.Button("Close Window", GUILayout.ExpandWidth(false), GUILayout.Width(100)))
                        Close();

                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();

                GUILayout.Space(10);
            }
            EditorGUILayout.EndVertical();
        }

        /// <summary>
        /// Draw cell table.
        /// </summary>
        private void DrawCellTable()
        {
            int click_idx = 0;
            click_idx = GUILayout.SelectionGrid(m_SelectIndex, m_CellTable.ToArray(), m_ColumnNumWithLabel, m_GUIStyleHolder.CellStyle,
                                                GUILayout.Width(m_CellWidth * m_ColumnNumWithLabel),
                                                GUILayout.Height(m_CellHeight * m_RowNumWithLabel));
            Rect rect_without_labels = new Rect(m_CellWidth, m_CellHeight, (m_ColumnNumWithLabel - 1) * m_CellWidth, (m_RowNumWithLabel - 1) * m_CellHeight);

            if ((Event.current.type == EventType.Used) &&
                (rect_without_labels.Contains(Event.current.mousePosition))) {
                if (Event.current.button == 0) { // Left click
                    m_SelectIndex = click_idx;

                    int start_row_idx = m_SelectIndex / m_ColumnNumWithLabel;
                    int start_col_idx = m_SelectIndex % m_ColumnNumWithLabel;
                    m_StartCell = new Vector2Int(start_col_idx, start_row_idx);
                    int start_rect_y = m_StartCell.y * m_CellHeight;
                    int start_rect_x = m_StartCell.x * m_CellWidth;
                    m_StartCellRect = new Rect(start_rect_x, start_rect_y, m_CellWidth, m_CellHeight);

                    m_ShowFrame = false;
                } else { // Right click
                    int last_row_idx = click_idx / m_ColumnNumWithLabel;
                    int last_col_idx = click_idx % m_ColumnNumWithLabel;
                    m_LastCell = new Vector2Int(last_col_idx, last_row_idx);

                    m_ShowFrame = true;
                }
            }
        }

        /// <summary>
        /// Color column and row labels.
        /// </summary>
        private void ColorLabels()
        {
            // Column Label
            for (int col_idx = 0; col_idx < m_ColumnNumWithLabel; col_idx++) {
                Rect column_label = new Rect(col_idx * m_CellWidth, 0, m_CellWidth, m_CellHeight);
                if (col_idx == 0)
                    GUI.Box(column_label, string.Empty, m_GUIStyleHolder.LabelStyle);
                else
                    GUI.Box(column_label, (col_idx - 1).ToString(), m_GUIStyleHolder.LabelStyle);
            }

            // Row Label
            for (int row_idx = 1; row_idx < m_RowNumWithLabel; row_idx++) {
                Rect row_label = new Rect(0, row_idx * m_CellHeight, m_CellWidth, m_CellHeight);
                GUI.Box(row_label, (row_idx - 1).ToString(), m_GUIStyleHolder.LabelStyle);
            }
        }

        /// <summary>
        /// Colors start cell and frame
        /// </summary>
        private void ColorCell()
        {
            if (m_ShowFrame) {
                int row_num = Mathf.Abs(m_StartCell.y - m_LastCell.y) + 1;
                int col_num = Mathf.Abs(m_StartCell.x - m_LastCell.x) + 1;

                int frame_rect_x = Mathf.Min(m_StartCell.x, m_LastCell.x) * m_CellWidth;
                int frame_rect_y = Mathf.Min(m_StartCell.y, m_LastCell.y) * m_CellHeight;

                // Only 1 cell is selected
                if ((row_num == 1) && (col_num == 1)) {
                    Rect frame_rect = new Rect(frame_rect_x, frame_rect_y, m_CellWidth, m_CellHeight);
                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameMCStyle);
                } else if ((row_num == 1) && (col_num > 1)) {
                    for (int col_idx = 0; col_idx < col_num; col_idx++) {
                        Rect frame_rect = new Rect(frame_rect_x + (col_idx * m_CellWidth), frame_rect_y, m_CellWidth, m_CellHeight);
                        if (col_idx == 0)
                            GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameRightSpaceStyle);
                        else if (col_idx == (col_num - 1))
                            GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameLeftSpaceStyle);
                        else
                            GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameHorizonStyle);
                    }
                } else if ((row_num > 1) && (col_num == 1)) {
                    for (int row_idx = 0; row_idx < row_num; row_idx++) {
                        Rect frame_rect = new Rect(frame_rect_x, frame_rect_y + (row_idx * m_CellHeight), m_CellWidth, m_CellHeight);
                        if (row_idx == 0)
                            GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameBottomSpaceStyle);
                        else if (row_idx == (row_num - 1))
                            GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameTopSpaceStyle);
                        else
                            GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameVerticalStyle);
                    }
                } else {
                    for (int row_idx = 0; row_idx < row_num; row_idx++) {
                        for (int col_idx = 0; col_idx < col_num; col_idx++) {
                            Rect frame_rect = new Rect(frame_rect_x + (col_idx * m_CellWidth),
                                                       frame_rect_y + (row_idx * m_CellHeight),
                                                       m_CellWidth, m_CellHeight);
                            if (row_idx == 0) {
                                if (col_idx == 0)
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameTLStyle);
                                else if (col_idx == (col_num - 1))
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameTRStyle);
                                else
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameTCStyle);
                            }  else if (row_idx == (row_num - 1)) { 
                                if (col_idx == 0)
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameBLStyle);
                                else if (col_idx == (col_num - 1))
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameBRStyle);
                                else
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameBCStyle);
                            } else {
                                if (col_idx == 0)
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameMLStyle);
                                else if (col_idx == (col_num - 1))
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameMRStyle);
                                else
                                    GUI.Box(frame_rect, "", m_GUIStyleHolder.FrameMCStyle);
                            }
                        }
                    }
                }

            }
            GUI.Box(m_StartCellRect, "", m_GUIStyleHolder.StartCellStyle);
        }
    }
}
