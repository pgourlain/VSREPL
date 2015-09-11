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
            }
            return pane;

        }
        static IVsOutputWindowPane BuildPane
        {
            get
            {
                IVsOutputWindowPane buildPane = GetPane(VSConstants.GUID_BuildOutputWindowPane, "Build");
                return buildPane;
            }
        }
        static IVsOutputWindowPane DebugPane
        {
            get
            {
                IVsOutputWindowPane buildPane = GetPane(VSConstants.GUID_OutWindowDebugPane, "Debug");
                return buildPane;
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
        internal static void LogDebug(string message)
        {
            DebugPane.OutputString(message);
        }

        internal static void LogDebugError(Exception ex)
        {
            DebugPane.OutputString(ex.ToString() + Environment.NewLine);
        }

        #endregion

        static void SaveAs(string oldFileName, ref string newScriptName, IComponentModel comp)
        {
            //make sure that fileName has .cs extension
            newScriptName = System.IO.Path.ChangeExtension(newScriptName, ".cs");
            //make sure that user enter only filename with any dicrectory
            newScriptName = System.IO.Path.GetFileName(newScriptName);

            oldFileName = System.IO.Path.Combine(VSTools.ScriptsDirectory, oldFileName);
            var newFileName = System.IO.Path.Combine(VSTools.ScriptsDirectory, newScriptName);

            var site = comp.GetService<IServiceProvider>();
            VsShellUtilities.RenameDocument(site, oldFileName, newFileName);
        }

        internal static void SaveAs(IWpfTextViewHost host, string oldScriptName, ref string fileName)
        {
            var comp = (Microsoft.VisualStudio.ComponentModelHost.IComponentModel)Package.GetGlobalService(typeof(Microsoft.VisualStudio.ComponentModelHost.SComponentModel));
            var svc = comp.GetService<IVsEditorAdaptersFactoryService>();
            var view = svc.GetViewAdapter(host.TextView);
            IVsTextLines vsTextLines;
            ErrorHandler.ThrowOnFailure(view.GetBuffer(out vsTextLines));
            IVsPersistDocData2 vsPersistDocData = (IVsPersistDocData2)vsTextLines;
            SaveAs(oldScriptName, ref fileName, comp);
            if (vsPersistDocData != null)
            {
                string newDoc;
                int saveCanceled;
                if (vsPersistDocData.SaveDocData(VSSAVEFLAGS.VSSAVE_Save, out newDoc, out saveCanceled) != VSConstants.S_OK)
                {
                    System.Diagnostics.Trace.WriteLine("error while saving C# script document");
                }
            }
        }

    }
}
