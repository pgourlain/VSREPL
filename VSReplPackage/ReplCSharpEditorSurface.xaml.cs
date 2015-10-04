using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Scripting.CSharp;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Text.Editor;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using VSReplPackage.Scripting;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Shell;
using System.Diagnostics;

namespace VSReplPackage
{
    /// <summary>
    /// Interaction logic for ReplCSharpEditorSurface.xaml 
    /// </summary>
    public partial class ReplCSharpEditorSurface : UserControl, INotifyPropertyChanged
    {
        IVsEditorAdaptersFactoryService _EditorAdapterFactory;
        IComponentModel _componentModel;
        internal IVsTextView _ViewAdapter;
        IWpfTextViewHost _TextViewHost;
        IVsWindowFrame _childWindowFrame;
        IVsWindowFrame parentWindowFrame;
        IVsPersistDocData _docData;

        bool _runningInProgress = true;
        public ReplCSharpEditorSurface()
        {
            _CurrentScriptName = System.IO.Path.GetFileName(VSTools.DefaultScriptFileName);
            InitializeComponent();
            this.DataContext = this;
            _componentModel = (IComponentModel)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SComponentModel));
            _EditorAdapterFactory = _componentModel.GetService<IVsEditorAdaptersFactoryService>();
        }

        #region script 

        private void OnScriptExecuted(object result, Exception ex)
        {
            if (this.Dispatcher.CheckAccess())
            {
                _runningInProgress = true;
            }
            else
            {
                this.Dispatcher.BeginInvoke(new Action<object, Exception>(OnScriptExecuted), DispatcherPriority.Normal, result, ex);
            }
        }

        string GetScriptText()
        {
            var host = _TextViewHost;
            if (host != null)
            {
                var scriptCode = host.TextView.TextBuffer.CurrentSnapshot.GetText();
                return scriptCode;
            }
            return @"Message(""Unable to retreive script from editor or script is empty."")";
        }

        private void ScriptSaveAs(object sender, RoutedEventArgs e)
        {
            //TODO: is an old ui, but it works...
            string fileName = Microsoft.VisualBasic.Interaction.InputBox("New script name", "Save script As", _CurrentScriptName, -1, -1);
            if (!string.IsNullOrEmpty(fileName))
            {
                try {
                    VSTools.SaveAs(_TextViewHost, _CurrentScriptName, ref fileName);
                    _ScriptNameEntries = null;
                    _CurrentScriptName = fileName;
                    NotifyPropertyChanged(string.Empty);
                }
                catch(Exception ex)
                {
                    VSTools.LogDebug("ScriptSaveAs Failed : {0}{1}", ex, Environment.NewLine);
                    VSTools.ActivateLogWindow();
                }
            }
        }
        private void LoadScript(string scriptName)
        {
            if (_childWindowFrame != null)
            {
                _childWindowFrame.CloseFrame((uint)__FRAMECLOSE.FRAMECLOSE_SaveIfDirty);
            }
            _TextViewHost = null;
            _ViewAdapter = null;
            _docData = null;
            HostCSharpEditor();
        }
        #endregion

        #region Properties
        private void RunCommandExecuted(object sender)
        {
            //best way is to use wpf binding, but right now
            _runningInProgress = false;
            var runner = new VSScriptRunner();
            runner.Run(GetScriptText(), OnScriptExecuted);
        }

        private bool RunCanExecute(object sender)
        {
            return _runningInProgress;
        }

        ICommand _RunCommand;
        public ICommand RunCommand
        {
            get
            {
                if (_RunCommand == null)
                {
                    _RunCommand = new Microsoft.VisualStudio.PlatformUI.DelegateCommand(RunCommandExecuted, RunCanExecute);
                }
                return _RunCommand;
            }
        }

        string _CurrentScriptName;

        public event PropertyChangedEventHandler PropertyChanged;

        public string CurrentScriptName
        {
            get
            {
                return _CurrentScriptName;
            }
            set
            {
                _CurrentScriptName = value;
                LoadScript(_CurrentScriptName);
                NotifyPropertyChanged("CurrentScriptName");
            }
        }

        ObservableCollection<string> _ScriptNameEntries = null;
        public ObservableCollection<string> ScriptNameEntries
        {
            get
            {
                if (_ScriptNameEntries == null)
                {
                    _ScriptNameEntries = new ObservableCollection<string>(
                                System.IO.Directory.GetFiles(VSTools.ScriptsDirectory, "*.cs")
                                .Select(x => System.IO.Path.GetFileName(x)));
                }
                return _ScriptNameEntries;
            }
        }

        private void NotifyPropertyChanged(string propName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propName));
            }
        }
        #endregion

        #region Hosting C# editor

        internal void OnToolWindowCreated(IVsWindowFrame windowFrame)
        {
            parentWindowFrame = windowFrame;
            HostCSharpEditor();
        }

        private void HostCSharpEditor()
        {
            if (_TextViewHost == null)
            {
                InitializeEditor(_CurrentScriptName);
            }
            if (_childWindowFrame != null)
            {
                //because VsShellUtilities.OpenDocumentWithSpecificEditor call Show, i should cause hide, before changing parent
                _childWindowFrame.Hide();
                var window = VsShellUtilities.GetWindowObject(parentWindowFrame);
                _childWindowFrame.SetProperty((int)__VSFPROPID2.VSFPROPID_ParentHwnd, window.HWnd);
                _childWindowFrame.SetProperty((int)__VSFPROPID2.VSFPROPID_ParentFrame, parentWindowFrame);
            }
        }

        // ----------------------------------------------------------------------------------
        /// <summary>
        /// Initialize the editor
        /// </summary>
        // ----------------------------------------------------------------------------------
        private void InitializeEditor(string scriptName)
        {

            Guid guid_microsoft_csharp_editor = new Guid("{A6C744A8-0E4A-4FC6-886A-064283054674}");
            Guid guid_microsoft_csharp_editor_with_encoding = new Guid("{08467b34-b90f-4d91-bdca-eb8c8cf3033a}");
            Guid editorType = guid_microsoft_csharp_editor;// VSConstants.VsEditorFactoryGuid.TextEditor_guid;
            Guid logicalViewGuid = Microsoft.VisualStudio.VSConstants.LOGVIEWID.Primary_guid;

            IVsWindowFrame ppWindowFrame = null;
            try
            {
                var psp = (System.IServiceProvider)ReplEditorPackage.CurrentPackage;
                var fileName = System.IO.Path.Combine(VSTools.ScriptsDirectory, scriptName);
                ppWindowFrame = Microsoft.VisualStudio.Shell.VsShellUtilities.OpenDocumentWithSpecificEditor(psp, fileName, editorType, logicalViewGuid);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("REPLException : " + ex.ToString());
                //System.Windows.MessageBox.Show(ex.ToString());
            }
            if (ppWindowFrame != null)
            {
                _ViewAdapter = VsShellUtilities.GetTextView(ppWindowFrame);
                _childWindowFrame = ppWindowFrame;
                _TextViewHost = _EditorAdapterFactory.GetWpfTextViewHost(_ViewAdapter);
                _docData = VSTools.GetPersistDocData(_TextViewHost);
                this.leftSide.Content = _TextViewHost;
            }
        }
        #endregion

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (e.Source == this)
            {
                e.CanExecute = IsDirtyDocument();
                e.Handled = true;
            }
            else
            {
                Trace.WriteLine("e.Source is not this");
            }
        }

        private bool IsDirtyDocument()
        {
            return VSTools.IsDirty(_docData);
        }

        private void CommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            VSTools.Save(_docData);
        }

        private void GotoSettings(object sender, RoutedEventArgs e)
        {
            ReplEditorPackage.CurrentPackage.ShowOptionPage();
        }
    }
}
