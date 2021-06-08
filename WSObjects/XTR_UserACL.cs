using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XTRWorkSite.WSObjects.Enums;

namespace XTRWorkSite.WSObjects
{
    public class XTR_UserACL
    {
        public XTR_EnumAccessRight Right
        {
            get
            {
                return (XTR_EnumAccessRight)(int)_WorkSiteOject.Right;
            }
        }



        public XTR_User User
        {
            get
            {
                XTR_User user = new XTR_User();
                user._WorkSiteOject = _WorkSiteOject.User;
                return user;
            }
        }

        internal IManage.IManUserACL _WorkSiteOject { get; set; }
    }
}
