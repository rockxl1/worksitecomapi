using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_Custom
    {
        public string CUSTOM_ALIAS
        {
            get { return _WorkSiteOject.Name; }
        }

        public string Description
        {
            get { return _WorkSiteOject.Description; }
        }

        public bool IsEnable
        {
            get { return _WorkSiteOject.Enabled; }

        }

        public XTR_Custom ParentAlias
        {
            get
            {
                XTR_Custom parent = new XTR_Custom();
                parent._WorkSiteOject = _WorkSiteOject.Parent;
                return parent;
            }
        }


        public List<XTR_Custom> ChillList
        {
            get
            {
                IManage.IManCustomAttributes chill = _WorkSiteOject.GetChildList(null, IManage.imSearchAttributeType.imSearchBoth, IManage.imSearchEnabledDisabled.imSearchEnabledOrDisabled, true);

                List<XTR_Custom> result = new List<XTR_Custom>();

                for (int f = 1; f <= chill.Count; f++)
                {
                    XTR_Custom custom = new XTR_Custom();
                    custom._WorkSiteOject = chill.ItemByIndex(f);
                    result.Add(custom);
                }
                return result;
            }
        }

        internal IManage.IManCustomAttribute _WorkSiteOject { get; set; }
    }
}
