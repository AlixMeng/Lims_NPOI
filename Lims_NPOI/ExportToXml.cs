using System;
using System.Xml;

namespace nsLims_NPOI
{
    class ExportToXml
    {
        private string _xml = "";

        private string _OutputPath = "";

        public string OutputPath
        {
            get
            {
                return this._OutputPath;
            }
            set
            {
                this._OutputPath = value;
            }
        }

        public string xml
        {
            get
            {
                return this._xml;
            }
            set
            {
                this._xml = value;
            }
        }

        public ExportToXml()
        {
        }

        public void createXMLFile(string _xml, string _OutputPath)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(_xml);
            xmlDocument.Save(_OutputPath);
        }
    }
}
