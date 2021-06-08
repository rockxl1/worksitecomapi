using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_DocumentType
    {
        public string ApplicationExtension
        {
            get { return _WorkSiteOject.ApplicationExtension; }
        }

        public string Description
        {
            get { return _WorkSiteOject.Description; }
        }

        public string ID
        {
            get { return _WorkSiteOject.ID; }
        }

        public string Name
        {
            get { return _WorkSiteOject.Name; }
        }

        public string AppExtension { get { return _WorkSiteOject.ApplicationExtension; } }

        internal IManage.IManDocumentType _WorkSiteOject { get; set; }

    }
}
