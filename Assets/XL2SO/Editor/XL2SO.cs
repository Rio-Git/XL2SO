using UnityEditor;
using System;
using System.Reflection;
using UnityEngine;
using System.Collections.Generic;

namespace XL2SO 
{
    /// <summary>
    /// Entry point of XL2SO extension.
    /// </summary>
    public class XL2SO : EditorWindow
    {
        private IScreen m_CurrentScreen = null; // Screen instance which run in current frame.
        private IScreen m_NextScreen = null;    // Screen instance which will run in next frame. 

        [MenuItem("Tools/XL2SO")]
        static void Open()
        {
            EditorWindow.GetWindow<XL2SO>("XL2SO");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="XL2SO"/> class.
        /// <remarks>
        /// Default screen is a instance of <see cref="Menu"/> class.
        /// </remarks>
        /// </summary>
        private void Awake()
        {
            m_CurrentScreen = CreateInstance<Menu>();
            m_CurrentScreen.Initialize(this);
            m_NextScreen = m_CurrentScreen;
        }

        /// <summary>
        /// OnGUI() function of Monobehaviour.
        /// </summary>
        /// <remarks>
        /// When screen is switched to another one:
        /// - Destroy current screen
        /// - Initialize new screen
        /// </remarks>
        private void OnGUI()
        {   
            m_NextScreen = m_CurrentScreen.Display();

            if (m_NextScreen != m_CurrentScreen)
            {
                DestroyImmediate(m_CurrentScreen);
                m_CurrentScreen = m_NextScreen;
                m_CurrentScreen.Initialize(this);
            }
        }
    }
}