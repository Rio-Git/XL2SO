using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace XL2SO
{
    /// <summary>
    /// Data class for GYUStyles.
    /// </summary>
    public class GUIStyleHolder : ScriptableObject
    {
        [SerializeField] private Texture2D CellTexture             = null; // [GUI] Texture for normal cell 
        [SerializeField] private Texture2D StartCellTexture        = null; // [GUI] Texture for start cell of selected range
        [SerializeField] private Texture2D LabelTexture            = null; // [GUI] Texture for labels
        [SerializeField] private Texture2D FrameTLTexture          = null; // [GUI] Texture for a frame at top left
        [SerializeField] private Texture2D FrameTCTexture          = null; // [GUI] Texture for a frame at top center
        [SerializeField] private Texture2D FrameTRTexture          = null; // [GUI] Texture for a frame at top right
        [SerializeField] private Texture2D FrameMLTexture          = null; // [GUI] Texture for a frame at middle left
        [SerializeField] private Texture2D FrameMCTexture          = null; // [GUI] Texture for a frame at middle center
        [SerializeField] private Texture2D FrameMRTexture          = null; // [GUI] Texture for a frame at middle right
        [SerializeField] private Texture2D FrameBLTexture          = null; // [GUI] Texture for a frame at bottom left
        [SerializeField] private Texture2D FrameBCTexture          = null; // [GUI] Texture for a frame at bottom center
        [SerializeField] private Texture2D FrameBRTexture          = null; // [GUI] Texture for a frame at bottom right
        [SerializeField] private Texture2D FrameTopSpaceTexture    = null; // [GUI] Texture for a frame which has a space at top
        [SerializeField] private Texture2D FrameBottomSpaceTexture = null; // [GUI] Texture for a frame which has a space at bottom
        [SerializeField] private Texture2D FrameRightSpaceTexture  = null; // [GUI] Texture for a frame which has a space at right
        [SerializeField] private Texture2D FrameLeftSpaceTexture   = null; // [GUI] Texture for a frame which has a space at left
        [SerializeField] private Texture2D FrameHorizonTexture     = null; // [GUI] Texture for a frame which has horizontal outlines
        [SerializeField] private Texture2D FrameVerticalTexture    = null; // [GUI] Texture for a frame which has vertical outlines
    
        public GUIStyle CellStyle             = null; // [GUI] GUIStyle for normal cell 
        public GUIStyle StartCellStyle        = null; // [GUI] GUIStyle for start cell of selected range
        public GUIStyle LabelStyle            = null; // [GUI] GUIStyle for labels
        public GUIStyle FrameTLStyle          = null; // [GUI] GUIStyle for a frame at top left
        public GUIStyle FrameTCStyle          = null; // [GUI] GUIStyle for a frame at top center
        public GUIStyle FrameTRStyle          = null; // [GUI] GUIStyle for a frame at top right
        public GUIStyle FrameMLStyle          = null; // [GUI] GUIStyle for a frame at middle left
        public GUIStyle FrameMCStyle          = null; // [GUI] GUIStyle for a frame at middle center
        public GUIStyle FrameMRStyle          = null; // [GUI] GUIStyle for a frame at middle right
        public GUIStyle FrameBLStyle          = null; // [GUI] GUIStyle for a frame at bottom left
        public GUIStyle FrameBCStyle          = null; // [GUI] GUIStyle for a frame at bottom center
        public GUIStyle FrameBRStyle          = null; // [GUI] GUIStyle for a frame at bottom right
        public GUIStyle FrameTopSpaceStyle    = null; // [GUI] GUIStyle for a frame which has a space at top
        public GUIStyle FrameBottomSpaceStyle = null; // [GUI] GUIStyle for a frame which has a space at bottom
        public GUIStyle FrameRightSpaceStyle  = null; // [GUI] GUIStyle for a frame which has a space at right
        public GUIStyle FrameLeftSpaceStyle   = null; // [GUI] GUIStyle for a frame which has a space at left
        public GUIStyle FrameHorizonStyle     = null; // [GUI] GUIStyle for a frame which has horizontal outlines
        public GUIStyle FrameVerticalStyle    = null; // [GUI] GUIStyle for a frame which has vertical outlines
    
        /// <summary>
        /// Initializes a new instance of <see cref="GUIStyleHolder"/> class.
        /// </summary>
        public void Initialize()
        {
            CellStyle = GetButtonStyle(CellTexture);
            CellStyle.padding = new RectOffset(2, 2, 2, 2);
    
            StartCellStyle = GetButtonStyle(StartCellTexture);
    
            LabelStyle = GetButtonStyle(LabelTexture);
            LabelStyle.alignment = TextAnchor.MiddleCenter;
    
            FrameTLStyle          = GetButtonStyle(FrameTLTexture);
            FrameTCStyle          = GetButtonStyle(FrameTCTexture);
            FrameTRStyle          = GetButtonStyle(FrameTRTexture);
            FrameMLStyle          = GetButtonStyle(FrameMLTexture);
            FrameMCStyle          = GetButtonStyle(FrameMCTexture);
            FrameMRStyle          = GetButtonStyle(FrameMRTexture);
            FrameBLStyle          = GetButtonStyle(FrameBLTexture);
            FrameBCStyle          = GetButtonStyle(FrameBCTexture);
            FrameBRStyle          = GetButtonStyle(FrameBRTexture);
            FrameTopSpaceStyle    = GetButtonStyle(FrameTopSpaceTexture);
            FrameBottomSpaceStyle = GetButtonStyle(FrameBottomSpaceTexture);
            FrameRightSpaceStyle  = GetButtonStyle(FrameRightSpaceTexture);
            FrameLeftSpaceStyle   = GetButtonStyle(FrameLeftSpaceTexture);
            FrameHorizonStyle     = GetButtonStyle(FrameHorizonTexture);
            FrameVerticalStyle    = GetButtonStyle(FrameVerticalTexture);
        }
    
        /// <summary>
        /// Gets a Button GUIStyle.
        /// </summary>
        /// <param name="_tex">Texture2D to be set as background image.</param>
        /// <returns>
        /// GUIStyle which has _tex background.
        /// </returns>
        private GUIStyle GetButtonStyle(Texture2D _tex)
        {
            GUIStyle style             =  new GUIStyle(GUI.skin.button);
            style.alignment            = TextAnchor.MiddleLeft;
            style.normal.textColor     = Color.black;
            style.hover.textColor      = Color.black;
            style.active.textColor     = Color.black;
            style.focused.textColor    = Color.black;
            style.normal.background    = _tex;
            style.active.background    = _tex;
            style.focused.background   = _tex;
            style.hover.background     = _tex;
            style.onNormal.textColor   = Color.black;
            style.onActive.textColor   = Color.black;
            style.onFocused.textColor  = Color.black;
            style.onHover.textColor    = Color.black;
            style.onNormal.background  = _tex;
            style.onActive.background  = _tex;
            style.onFocused.background = _tex;
            style.onHover.background   = _tex;
            style.margin               = new RectOffset(0, 0, 0, 0);
            style.border               = new RectOffset(0, 0, 0, 0);
            style.padding              = new RectOffset(0, 0, 0, 0);
    
            return style;
        }
    }
}
