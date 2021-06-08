using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_DocumentClass
    {
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

        internal IManage.IManDocumentClass _WorkSiteOject { get; set; }

    }
}
