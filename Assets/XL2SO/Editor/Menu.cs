using UnityEngine;
using UnityEditor;

namespace XL2SO 
{
    /// <summary>
    /// Provides a Menu screen.
    /// </summary>
    public class Menu : IScreen
    {
        public Texture2D m_IG_Texture = null;
        public Texture2D m_SG_Texture = null;

        /// <summary>Main routine of Menu.</summary>
        /// <returns>
        /// SG instance: When SG button is clicked.
        /// IG instance: When IG button is clicked.
        /// Menu instance (this): Other than above cases.
        /// </returns>
        override public IScreen Display()
        {
            bool IG_clicked = false;
            bool SG_clicked = false;

            EditorGUILayout.BeginVertical();
            {
                GUILayout.FlexibleSpace();

                // SG
                EditorGUILayout.BeginHorizontal();
                {
                    GUILayout.FlexibleSpace();
                    SG_clicked = GUILayout.Button(m_SG_Texture, GUIStyle.none);
                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();

                EditorGUILayout.BeginHorizontal();
                {
                    GUILayout.FlexibleSpace();
                    EditorGUILayout.LabelField("Generate ScriptableObject Script", EditorStyles.boldLabel, GUILayout.Width(220));
                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();

                GUILayout.FlexibleSpace();
                
                // IG
                EditorGUILayout.BeginHorizontal();
                {
                    GUILayout.FlexibleSpace();
                    IG_clicked = GUILayout.Button(m_IG_Texture, GUIStyle.none);
                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();
                EditorGUILayout.BeginHorizontal();
                {
                    GUILayout.FlexibleSpace();
                    EditorGUILayout.LabelField("Generate ScriptableObject Instances", EditorStyles.boldLabel, GUILayout.Width(250));
                    GUILayout.FlexibleSpace();
                }
                EditorGUILayout.EndHorizontal();

                GUILayout.FlexibleSpace();
            }
            EditorGUILayout.EndVertical();

            if (SG_clicked)
                return CreateInstance<SG.SG>();
            else if (IG_clicked)
                return CreateInstance<IG.IG>();
            else
                return this;

        }
    }

}