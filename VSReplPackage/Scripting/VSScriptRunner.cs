﻿using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Scripting.CSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSReplPackage.Scripting
{
    class VSScriptRunner
    {
        static string[] defaultReferences = { "System", "system.core", "mscorlib", "Microsoft.CSharp", "EnvDte", "EnvDTE100",
        "Microsoft.VisualStudio.Shell.14.0",
        "Microsoft.VisualStudio.Shell.Interop",
        "Microsoft.VisualStudio.Shell.Interop.10.0",
        "Microsoft.VisualStudio.Shell.Interop.11.0",
        "Microsoft.VisualStudio.Shell.Interop.12.0",
        "Microsoft.VisualStudio.Shell.Interop.8.0",
        "Microsoft.VisualStudio.Shell.Interop.9.0",
        };
        static string[] defaultNamespaces = { "System", "System.Linq", "System.Text", "System.Collections.Generic",
            "System.Diagnostics", "System.IO", "EnvDTE",
            "Microsoft.VisualStudio","Microsoft.VisualStudio.Shell", "Microsoft.VisualStudio.Shell.Interop"
        };

        static ScriptOptions defaultOptions;
        static VSScriptRunner()
        {
            var VsDir = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
            Trace.WriteLine("VSDir : " + VsDir);
            //TODO: find location of Visual Studio installation
            //location of package
            var path = System.IO.Path.GetDirectoryName(typeof(CSharpScript).Assembly.Location);
            var options = ScriptOptions.Default.WithBaseDirectory(path)
                .WithSearchPaths(
                @"c:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\PublicAssemblies",
                @"c:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\PrivateAssemblies");
            options = options.AddReferences(typeof(CSharpScript).Assembly,
                typeof(EnvDTE.TextSelection).Assembly,
                typeof(EnvDTE.DTE).Assembly);
            options = options.WithReferences(defaultReferences);
            options = options.WithNamespaces(defaultNamespaces);
            defaultOptions = options;
        }

        public void Run(string sCode, Action<object, Exception> executionEnd)
        {
            VSTools.LogStartRunning();
            try
            {
                var result = CSharpScript.RunAsync(sCode, defaultOptions, globals: new Globals());

                result.ContinueWith(OnException, executionEnd, TaskContinuationOptions.OnlyOnFaulted);
                result.ContinueWith(OnSuccess, executionEnd, TaskContinuationOptions.OnlyOnRanToCompletion);
            }
            catch(Exception ex)
            {
                VSTools.LogError(ex);
                VSTools.LogEndRunning(true);
                executionEnd(null, ex);
            }
        }
        private void OnException(Task<ScriptState<object>> t, object state)
        {
            var ex = t.Exception;
            VSTools.LogError(ex);
            VSTools.LogEndRunning(true);
            var executionEnd = (Action<object, Exception>)state;
            executionEnd(null, ex);
        }

        private void OnSuccess(Task<ScriptState<object>> t, object state)
        {
            var result = t.IsCompleted ? t.Result : null;
            var exception = t.IsFaulted ? t.Exception : null;
            VSTools.LogEndRunning(t.IsFaulted);
            var executionEnd = (Action<object, Exception>)state;
            executionEnd(result, exception);
        }

    }
}
