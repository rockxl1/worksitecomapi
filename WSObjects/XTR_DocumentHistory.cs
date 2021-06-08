using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_DocumentHistory
    {
        public string Application
        {
            get { return _WorkSiteOject.Application; }
        }

        public string Comment
        {
            get { return _WorkSiteOject.Comment; }
        }

        public DateTime Date
        {
            get { return _WorkSiteOject.Date; }
        }

        public int Duration
        {
            get { return _WorkSiteOject.Duration; }
        }

        public string Location
        {
            get { return _WorkSiteOject.Location; }
        }

        public string Number
        {
            get { return _WorkSiteOject.Number; }
        }

        public string Operation
        {
            get { return _WorkSiteOject.Operation; }
        }

        public int PagesPrinted
        {
            get { return _WorkSiteOject.PagesPrinted; }
        }

        public string User
        {
            get { return _WorkSiteOject.User; }
        }

        public object UserNumberProperty1
        {
            get { return _WorkSiteOject.UserNumberProperty1; }
        }

        public object UserNumberProperty2
        {
            get { return _WorkSiteOject.UserNumberProperty2; }
        }

        public object UserNumberProperty3
        {
            get { return _WorkSiteOject.UserNumberProperty3; }
        }

        public string UserProperty1
        {
            get { return _WorkSiteOject.UserProperty1; }
        }

        public string UserProperty2
        {
            get { return _WorkSiteOject.UserProperty2; }
        }

        public string Version
        {
            get { return _WorkSiteOject.Version; }
        }

        internal IManage.IManDocumentHistory _WorkSiteOject { get; set; }
    }
}
