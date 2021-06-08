using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_User
    {
        public string FullName
        {
            get { return _WorkSiteOject.FullName; }
        }

        public bool isLoginEnable
        {
            get { return _WorkSiteOject.LoginEnabled; }
        }

        public string Email
        {
            get { return _WorkSiteOject.Email; }
        }

        public string Name
        {
            get
            {
                return _WorkSiteOject.Name;
            }
        }

        public string Custom1
        {
            get
            {
                return _WorkSiteOject.Custom1;
            }
        }

        public List<XTR_Group> Groups
        {
            get
            {
                List<XTR_Group> groupAcl = new List<XTR_Group>();

                if (_WorkSiteOject.Groups != null && _WorkSiteOject.Groups.Count > 0)
                {
                    for (int i = 1; i <= _WorkSiteOject.Groups.Count; i++)
                    {
                        XTR_Group item = new XTR_Group();
                        item._WorkSiteOject = _WorkSiteOject.Groups.ItemByIndex(i);
                        groupAcl.Add(item);
                    }
                }

                return groupAcl;
            }
        }

        internal IManage.IManUser _WorkSiteOject { get; set; }
    }
}
