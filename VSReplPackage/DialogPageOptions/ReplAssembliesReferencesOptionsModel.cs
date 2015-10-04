using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace VSReplPackage
{

    [Serializable]
    public class ReplAssembliesReferencesOptionsModel : INotifyPropertyChanged
    {
        private string[] _SearchPaths;
        private string[] _References;
        private string[] _Namespaces;
        private string _CustomSettings;

        public ReplAssembliesReferencesOptionsModel()
        {
            _CustomSettings = string.Empty;
            this.SearchPaths = new string[0];
            this.References = new string[0];
            this.Namespaces = new string[0];
        }
        #region storage management
        public void UpdatePropertiesFromStorage()
        {
            this._References = DeSerializeStringArray(this.ReferencesStorage);
            this._SearchPaths = DeSerializeStringArray(this.SearchPathsStorage);
            this._Namespaces = DeSerializeStringArray(this.NamespacesStorage);
        }

        private void UpdateStorages()
        {
            this.ReferencesStorage = SerializeStringArray(this._References);
            this.SearchPathsStorage = SerializeStringArray(this._SearchPaths);
            this.NamespacesStorage = SerializeStringArray(this._Namespaces);
        }

        private string SerializeStringArray(string[] array)
        {
            if (array != null)
            {
                return string.Join(";", array);
            }
            return string.Empty;
        }
        private string[] DeSerializeStringArray(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                return value.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            }
            return new string[0];
        }

        /// <summary>
        /// just to avoid a custom serializer for string[]
        /// </summary>
        [Browsable(false)]
        public string ReferencesStorage { get; set; }

        /// <summary>
        /// just to avoid a custom serializer for string[]
        /// </summary>
        [Browsable(false)]
        public string NamespacesStorage { get; set; }

        /// <summary>
        /// just to avoid a custom serializer for string[]
        /// </summary>
        [Browsable(false)]
        public string SearchPathsStorage { get; set; }
        #endregion


        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [XmlArrayAttribute]
        [Category("Script search paths")]
        [Description("Add search paths. Some Visual Studio paths(PublicAssemblies/PrivateAssemblies directories) are automatically added.")]
        [DisplayName("SearchPath")]
        public string[] SearchPaths { get { return _SearchPaths; } set { _SearchPaths = value; UpdateStorages(); } }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [XmlArrayAttribute]
        [Category("Script setup")]
        [Description("Assemblies that should be used in script. Add filename without extension.")]
        [DisplayName("References")]
        public string[] References { get { return _References; } set { _References = value; UpdateStorages(); } }


        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [XmlArrayAttribute]
        [Category("Script setup")]
        [Description("'using namespace' that you want in your scripts.")]
        [DisplayName("Namespaces")]
        public string[] Namespaces { get { return _Namespaces; } set { _Namespaces = value; UpdateStorages(); } }

        [Category("Custom templates file")]
        [Description("File that contains your custom setting templates")]
        [DisplayName("CustomSettings")]
        public string CustomSettings { get { return _CustomSettings; }
            set
            {
                _CustomSettings = value;
                DoPropertyChanged("CustomSettings");
                this.CustomTemplates = null;
            }
        }


        ObservableCollection<string> _CustomTemplates;
        [Browsable(false)]
        public ObservableCollection<string> CustomTemplates
        {
            get
            {
                if (_CustomTemplates == null)
                {
                    _CustomTemplates = new ObservableCollection<string>();
                    FillTemplates();
                }
                return _CustomTemplates;
            }
            set
            {
                _CustomTemplates = value;
                DoPropertyChanged("CustomTemplates");
            }
        }

        internal void ApplyCustomTemplate(string templateName)
        {
            XDocument doc = XDocument.Load(this.CustomSettings);
            var setting = doc.XPathSelectElement("/Document/Setting[@name='" + templateName + "']");
            if (setting != null)
            {
                this.SearchPathsStorage = GetXmlText(setting, "SearchPaths");
                this.ReferencesStorage = GetXmlText(setting, "References");
                this.NamespacesStorage = GetXmlText(setting, "Namespaces");
            }
            UpdatePropertiesFromStorage();
        }

        private string GetXmlText(XElement setting, string nodeName)
        {
            var xelem = setting.XPathSelectElement(nodeName);
            if (xelem != null)
            {
                return xelem.Value;
            }
            return string.Empty;
        }

        private void FillTemplates()
        {
            if (File.Exists(this.CustomSettings))
            {
                try
                {
                    XDocument doc = XDocument.Load(this.CustomSettings);
                    foreach (var element in doc.XPathSelectElements("/Document/Setting"))
                    {
                        var attrName = element.Attribute("name");
                        if (attrName != null)
                        {
                            _CustomTemplates.Add(attrName.Value);
                        }
                    }
                }
                catch
                {
                    //ignore all exception while loading xml
                }
            }
        }

        internal void AllPropertiesChanged()
        {
            DoPropertyChanged(string.Empty);
        }

        private void DoPropertyChanged(string propName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
