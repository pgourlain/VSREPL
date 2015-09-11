//------------------------------------------------------------------------------
// <copyright file="ReplEditor.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace VSReplPackage
{
    using System;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.ComponentModelHost;
    using Microsoft.VisualStudio.Editor;
    using Microsoft.VisualStudio.OLE.Interop;
    using Microsoft.VisualStudio.Shell.Interop;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Text;
    using Microsoft.VisualStudio.Text.Editor;
    using Microsoft.VisualStudio.TextManager.Interop;
    using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
    using Microsoft.VisualStudio.Utilities;
    using System.Linq;
    using System.IO;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Text.Projection;
    using System.Windows;
    using System.Windows.Input;
    using System.Text;
    using System.Reflection;
    using Scripting;

    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("4292b9da-ff15-460b-92d8-b4e964ce29ec")]
    public class ReplEditor : ToolWindowPane, IOleCommandTarget
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="ReplEditor"/> class.
        /// </summary>
        public ReplEditor() : base(null)
        {
            this.Caption = "Repl Script Editor for Visual Studio - " + typeof(ReplEditor).Assembly.GetName().Version.ToString();
            //deploy sample script if needed
            var di = new DirectoryInfo(VSTools.ScriptsDirectory);
            if (!di.Exists) di.Create();

            var fileName = VSTools.DefaultScriptFileName;
            if (!File.Exists(fileName))
            {
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "VSReplPackage.SampleScript.cs";

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                using (StreamReader reader = new StreamReader(stream))
                {
                    string s = reader.ReadToEnd();
                    File.WriteAllText(fileName, s, Encoding.UTF8);
                }
            }
            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            //this.Content = new ReplEditorControl();
            //init services
        }

        ReplCSharpEditorSurface _surface = null;

        // ----------------------------------------------------------------------------------
        /// <summary>
        /// Get the content of the window pane: an instance of the editor to be embedded into
        /// the pane.
        /// </summary>
        // ----------------------------------------------------------------------------------
        override public object Content
        {
            get
            {
                if (_surface == null)
                {
                    _surface = new ReplCSharpEditorSurface();
                }
                return _surface;
            }
        }



        // ----------------------------------------------------------------------------------
        /// <summary>
        /// Gets the editor wpf host that we can use as the tool windows content.
        /// </summary>
        // ----------------------------------------------------------------------------------
        //public IWpfTextViewHost TextViewHost
        //{
        //    get
        //    {
        //        if (_TextViewHost == null)
        //        {
        //            InitializeEditor();
        //            _TextViewHost = _EditorAdapterFactory.GetWpfTextViewHost(_ViewAdapter);
        //            /*
        //            var data = _ViewAdapter as IVsUserData;
        //            if (data != null)
        //            {
        //                var guid = Microsoft.VisualStudio.Editor.DefGuidList.guidIWpfTextViewHost;
        //                object obj;
        //                var hr = data.GetData(ref guid, out obj);
        //                if ((hr == Microsoft.VisualStudio.VSConstants.S_OK) &&
        //                  obj != null && obj is IWpfTextViewHost)
        //                {
        //                    _TextViewHost = obj as IWpfTextViewHost;
        //                }
        //            }
        //            */
        //        }
        //        return _TextViewHost;
        //    }
        //}

        // ----------------------------------------------------------------------------------
        /// <summary>
        /// Register key bindings to use in the editor.
        /// </summary>
        // ----------------------------------------------------------------------------------
        public override void OnToolWindowCreated()
        {
            // --- Register key bindings to use in the editor
            var windowFrame = (IVsWindowFrame)Frame;
            var cmdUi = Microsoft.VisualStudio.VSConstants.GUID_TextEditorFactory;
            windowFrame.SetGuidProperty((int)__VSFPROPID.VSFPROPID_InheritKeyBindings, ref cmdUi);
            _surface.OnToolWindowCreated(windowFrame);
            base.OnToolWindowCreated();

        }

        const int WM_SETFOCUS = 0x0007;
        const int WM_KEYDOWN = 0x0100;
        const int WM_KEYUP = 0x0101;

        // ----------------------------------------------------------------------------------
        /// <summary>
        /// Allow the embedded editor to handle keyboard messages before they are dispatched
        /// to the window that has the focus.
        /// </summary>
        // ----------------------------------------------------------------------------------
        protected override bool PreProcessMessage(ref Message m)
        {
            if (_surface != null && _surface._ViewAdapter != null)
            {
                // copy the Message into a MSG[] array, so we can pass
                // it along to the active core editor's IVsWindowPane.TranslateAccelerator
                var pMsg = new MSG[1];
                pMsg[0].hwnd = m.HWnd;
                pMsg[0].message = (uint)m.Msg;
                pMsg[0].wParam = m.WParam;
                pMsg[0].lParam = m.LParam;

                var vsWindowPane = (IVsWindowPane)_surface._ViewAdapter;
                return vsWindowPane.TranslateAccelerator(pMsg) == 0;
            }
            return base.PreProcessMessage(ref m);
        }

        // ----------------------------------------------------------------------------------
        /// <summary>
        /// Forwards command execution messages recevived by the window pane to the embedded
        /// editor.
        /// </summary>
        // ----------------------------------------------------------------------------------
        int IOleCommandTarget.Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt,
          IntPtr pvaIn, IntPtr pvaOut)
        {
            var hr = (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
            if (_surface != null)
            {
                var cmdTarget = (IOleCommandTarget)_surface._ViewAdapter;
                hr = cmdTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }
            return hr;
        }

        // ----------------------------------------------------------------------------------
        /// <summary>
        /// Forwards command status query messages recevived by the window pane to the 
        /// embedded editor.
        /// </summary>
        // ----------------------------------------------------------------------------------
        int IOleCommandTarget.QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds,
          IntPtr pCmdText)
        {
            var hr = (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
            if (_surface != null && _surface._ViewAdapter != null)
            {
                var cmdTarget = (IOleCommandTarget)_surface._ViewAdapter;
                hr = cmdTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
            }
            return hr;
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }
    }
}
