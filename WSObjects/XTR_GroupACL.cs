using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XTRWorkSite.WSObjects.Enums;

namespace XTRWorkSite.WSObjects
{
    public class XTR_GroupACL
    {

        public XTR_EnumAccessRight Right
        {
            get { return (XTR_EnumAccessRight)(int)_WorkSiteOject.Right; }
        }

        public XTR_Group Group
        {
            get
            {
                XTR_Group group = new XTR_Group();
                group._WorkSiteOject = _WorkSiteOject.Group;
                return group;

            }
        }

        internal IManage.IManGroupACL _WorkSiteOject { get; set; }

    }
}
