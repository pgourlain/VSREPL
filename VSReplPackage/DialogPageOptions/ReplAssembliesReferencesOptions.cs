using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSReplPackage
{
    //TODO: inherits from UIElementDialogPage in order to use WPF in option page
    [Guid("FA4E59E6-3695-4FD4-9CE4-254E2BF8923F")]
    [ComVisible(true)]
    public sealed class ReplAssembliesReferencesOptions : UIElementDialogPage
    {
        private ReplAssembliesReferencesOptionsModel _model = new ReplAssembliesReferencesOptionsModel();

        private AssembliesReferencesOptionsUI _ui;

        public ReplAssembliesReferencesOptions()
        {
            _ui = new AssembliesReferencesOptionsUI();
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            base.OnApply(e);
            Scripting.VSScriptRunner.UpdateDefaultOptions(_model);
        }

        public override object AutomationObject
        {
            get
            {
                return _model;
            }
        }

        protected override UIElement Child
        {
            get
            {
                return _ui;
            }
        }

        public override void LoadSettingsFromStorage()
        {
            base.LoadSettingsFromStorage();
            var model = this.AutomationObject as ReplAssembliesReferencesOptionsModel;
            if (model != null)
            {
                model.UpdatePropertiesFromStorage();
            }
            _ui.AutomationObject = model;
        }

        public override void SaveSettingsToStorage()
        {
            base.SaveSettingsToStorage();
        }
    }
}
