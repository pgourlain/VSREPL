using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
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

namespace VSReplPackage
{
    /// <summary>
    /// Interaction logic for AssembliesReferencesOptionsUI.xaml
    /// </summary>
    public partial class AssembliesReferencesOptionsUI : UserControl, INotifyPropertyChanged
    {
        System.Windows.Forms.PropertyGrid _pg;
        public AssembliesReferencesOptionsUI()
        {
            InitializeComponent();
            _pg = new System.Windows.Forms.PropertyGrid();
            this.win32Host.Child = _pg;
            this.DataContext = this;
        }

        ICommand _CustomTemplateCommand;

        public event PropertyChangedEventHandler PropertyChanged;
        private void DoNotifyPropertyChanged(string propName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }

        public ICommand CustomTemplateCommand
        {
            get
            {
                if (_CustomTemplateCommand == null)
                {
                    _CustomTemplateCommand = new Microsoft.VisualStudio.PlatformUI.DelegateCommand(TemplateItem_Click);
                }
                return _CustomTemplateCommand;
            }
        }

        public ReplAssembliesReferencesOptionsModel AutomationObject
        {
            get
            {
                return (ReplAssembliesReferencesOptionsModel)_pg.SelectedObject;
            }
            set
            {
                _pg.SelectedObject = value;
                DoNotifyPropertyChanged("AutomationObject");
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void CodeFixProvider_Click(object sender, RoutedEventArgs e)
        {
            //add codefix settings
            this.AutomationObject.References = new string[] {"Microsoft.CodeAnalysis", "Microsoft.CodeAnalysis.CSharp",
                "Microsoft.CodeAnalysis.CSharp.Workspaces", "Microsoft.CodeAnalysis.Workspaces", "System.Collections.Immutable", "System.Composition.AttributedModel",
                "System.Composition.Convention", "System.Composition.Hosting", "System.Composition.Runtime", "System.Composition.TypedParts", "System.Reflection.Metadata" };
            this.AutomationObject.Namespaces = new string[] {
            "System.Linq",
            "System.Threading",
            "Microsoft.CodeAnalysis",
            "Microsoft.CodeAnalysis.CSharp",
            "Microsoft.CodeAnalysis.CSharp.Syntax",
            "Microsoft.CodeAnalysis.Diagnostics"};
        }

        private void Default_click(object sender, RoutedEventArgs e)
        {
            this.AutomationObject.Namespaces = Scripting.VSScriptRunner.defaultNamespaces;
            this.AutomationObject.References = Scripting.VSScriptRunner.defaultReferences;
            this.AutomationObject.AllPropertiesChanged();
            //add default;
        }

        private void CopyXmlTemplate_click(object sender, RoutedEventArgs e)
        {
            var xmlSample = @"<Document>
    <Setting name=""sample"">
        <SearchPaths>c:\tmp;c:\framework1;</SearchPaths>
        <References>" + string.Join(";", Scripting.VSScriptRunner.defaultReferences) + @"</References>
        <Namespaces>" + string.Join(";", Scripting.VSScriptRunner.defaultNamespaces) + @"</Namespaces>
    </Setting>
</Document>";
            Clipboard.SetText(xmlSample);
        }

        private void CustomTemplates_Click(object sender, RoutedEventArgs e)
        {
            Button senderButton = (sender as Button);
            senderButton.ContextMenu.Items.Clear();
            senderButton.ContextMenu.IsEnabled = true;
            senderButton.ContextMenu.PlacementTarget = senderButton;
            senderButton.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            foreach (var item in this.AutomationObject.CustomTemplates)
            {
                senderButton.ContextMenu.Items.Add(new MenuItem() { Header = item, Command = this.CustomTemplateCommand, CommandParameter=item });
            }
            senderButton.ContextMenu.IsOpen = true;
        }

        private void TemplateItem_Click(object arg)
        {
            this.AutomationObject.ApplyCustomTemplate(arg as string);
        }
    }
}
