using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace VSReplPackage.Scripting
{
    /// <summary>
    /// globals vars and methods for script engine
    /// </summary>
    public class Globals
    {
        public Globals()
        {
            this.DTE = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));
        }

        #region Variables
        public EnvDTE.DTE DTE;
        #endregion

        //if the script ends with a call to a void method, an error occured. So in order to avoid this error, my methods always return an int
        #region APIs
        public int MessageBox(string m)
        {
            System.Windows.MessageBox.Show(m);
            return 0;
        }

        public int Trace(string message)
        {
            VSTools.LogDebug(message);
            return 0;
        }
        public int TraceError(Exception ex)
        {
            VSTools.LogDebugError(ex);
            return 0;
        }
        #endregion

    }
}
