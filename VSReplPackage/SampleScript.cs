/*
Author : Pierrick Gourlain
script location : My Documents\Visual Studio 2015\Scripts\ReplScriptEditor.cs 
    help : if you delete this file, it will be re-created after restarting Visual Studio
-- Assemblies References --
"System", "system.core", "mscorlib", "Microsoft.CSharp", "EnvDte", "EnvDTE100", "Microsoft.VisualStudio.Shell.14.0", "Microsoft.VisualStudio.Shell.Interop.*"
-- Predefined namespaces --
"System.Linq", "System.Text", "System.Collections.Generic", "System.Diagnostics", "System.IO", "EnvDTE"

-- Globals variables --
DTE
-- Globals Methods --
MessageBox(string) : show MessageBox
Trace(string) : trace message to DebugPaneWindow
TraceError(Exception) : trace exception to DebugPaneWindow
*/



dynamic dte = DTE.ActiveDocument;
MessageBox(dte.Selection.Text);