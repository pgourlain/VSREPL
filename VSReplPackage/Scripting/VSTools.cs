using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.ComponentModelHost;

namespace VSReplPackage.Scripting
{
    class VSTools
    {
        static Guid GUID_ScriptEditorOutputWindowPane = new Guid("8ED8DFF5-EAAC-4222-B5AD-FECD608562BC");
        #region folder section
        public static string ScriptsDirectory
        {
            get
            {
                var VS2015ScriptsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Visual Studio 2015", "Scripts");
                return VS2015ScriptsDir;
            }
        }

        public static string DefaultScriptFileName
        {
            get
            {
                return Path.Combine(ScriptsDirectory, DefaultScriptName);
            }
        }

        /// <summary>
        /// name of default script name, without any directory
        /// </summary>
        public static string DefaultScriptName
        {
            get
            {
                return "ReplScriptEditor.cs";
            }
        }
        #endregion

        #region Logging section

        static IVsOutputWindowPane GetPane(Guid guidPane, string defaultText)
        {
            IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            IVsOutputWindowPane pane;
            if (outWindow.GetPane(ref guidPane, out pane) != VSConstants.S_OK)
            {
                outWindow.CreatePane(ref guidPane, defaultText, 1, 1);
                outWindow.GetPane(ref guidPane, out pane);
                pane.Activate();
            }
            return pane;

        }
        static IVsOutputWindowPane BuildPane
        {
            get
            {
                IVsOutputWindowPane buildPane = GetPane(GUID_ScriptEditorOutputWindowPane, "Repl Script Editor");
                return buildPane;
            }
        }
        static IVsOutputWindowPane DebugPane
        {
            get
            {               
                return BuildPane;
            }
        }

        internal static void LogError(Exception ex)
        {
            IVsOutputWindowPane buildPane = BuildPane;
            buildPane.OutputString("  Roslyn script Error:" + Environment.NewLine);
            buildPane.OutputString(ex.ToString() + Environment.NewLine);
        }

        internal static void LogStartRunning()
        {
            IVsOutputWindowPane buildPane = BuildPane;
            buildPane.Clear();
            buildPane.OutputString(String.Format(CultureInfo.InvariantCulture, "------ Roslyn script started: {0} ------{1}", DateTime.Now, Environment.NewLine));
        }

        internal static void LogEndRunning(bool failed)
        {
            IVsOutputWindowPane buildPane = BuildPane;
            buildPane.OutputString(string.Format(CultureInfo.InvariantCulture, "========== Roslyn script: execution {0}: {1} =========={2}", failed ? "failed" : "succeeded", DateTime.Now, Environment.NewLine));
            buildPane.Activate(); // Brings this pane into view
        }

        internal static void ActivateLogWindow()
        {
            BuildPane.Activate();
        }

        internal static void LogDebug(string message)
        {
            BuildPane.OutputString(message);
        }
        internal static void LogDebug(string fmt, params object[] args)
        {
            BuildPane.OutputString(string.Format(CultureInfo.InvariantCulture, fmt, args));
        }

        internal static void LogDebugError(Exception ex)
        {
            BuildPane.OutputString(ex.ToString() + Environment.NewLine);
        }

        #endregion

        static void SaveAs(string oldFileName, ref string newScriptName, IServiceProvider site)
        {
            //make sure that fileName has .cs extension
            newScriptName = System.IO.Path.ChangeExtension(newScriptName, ".cs");
            //make sure that user enter only filename with any dicrectory
            newScriptName = System.IO.Path.GetFileName(newScriptName);

            oldFileName = System.IO.Path.Combine(VSTools.ScriptsDirectory, oldFileName);
            var newFileName = System.IO.Path.Combine(VSTools.ScriptsDirectory, newScriptName);

            VsShellUtilities.RenameDocument(site, oldFileName, newFileName);
        }

        internal static void SaveAs(IWpfTextViewHost host, string oldScriptName, ref string fileName)
        {
            IVsPersistDocData vsPersistDocData = GetPersistDocData(host);
            SaveAs(oldScriptName, ref fileName, ServiceProvider.GlobalProvider);             
            if (vsPersistDocData != null)
            {
                Save(vsPersistDocData);
            }
        }

        private static IVsEditorAdaptersFactoryService _EditorAdaptersFactoryService;
        private static IVsEditorAdaptersFactoryService EditorAdaptersFactoryService
        {
            get
            {
                if (_EditorAdaptersFactoryService == null)
                {
                    var comp = (Microsoft.VisualStudio.ComponentModelHost.IComponentModel)Package.GetGlobalService(typeof(Microsoft.VisualStudio.ComponentModelHost.SComponentModel));
                    _EditorAdaptersFactoryService = comp.GetService<IVsEditorAdaptersFactoryService>();
                }
                return _EditorAdaptersFactoryService;
            }
        }

        internal static IVsPersistDocData GetPersistDocData(IWpfTextViewHost host)
        {
            var svc = EditorAdaptersFactoryService;
            if (svc == null) LogDebug("unable to get service 'EditorAdaptersFactoryService'");
            var view = svc.GetViewAdapter(host.TextView);
            IVsTextLines vsTextLines;
            ErrorHandler.ThrowOnFailure(view.GetBuffer(out vsTextLines));
            IVsPersistDocData vsPersistDocData = (IVsPersistDocData)vsTextLines;
            return vsPersistDocData;
        }

        internal static bool IsDirty(IVsPersistDocData docData)
        {
            int pfDirty;
            ErrorHandler.ThrowOnFailure(docData.IsDocDataDirty(out pfDirty));
            return pfDirty > 0;
        }

        internal static void Save(IVsPersistDocData docData)
        {
            int canceled;
            string newDocumentStateScope;
            if (docData.SaveDocData(VSSAVEFLAGS.VSSAVE_Save, out newDocumentStateScope, out canceled) != VSConstants.S_OK)
            {
                LogDebug("error while saving C# script document");
            }
        }
    }
}
