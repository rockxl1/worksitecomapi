using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_Group
    {
        public string Name {
            get {
                return _WorkSiteOject.Name;
            }
        }

        public List<XTR_User> Users
        {
            get
            {
                List<XTR_User> usrs = new List<XTR_User>();
                foreach (IManage.IManUser item in _WorkSiteOject.Users)
                {
                    XTR_User cursor = new XTR_User();
                    cursor._WorkSiteOject = item;
                    usrs.Add(cursor);
                }

                return usrs;
            }
        }

        internal IManage.IManGroup _WorkSiteOject { get; set; }

    }
}
