using IManage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XTRWorkSite.WSObjects.Enums;

namespace XTRWorkSite.WSObjects
{
    public class XTR_Security
    {
        public XTR_EnumImSecurityType DefaultOrSharedAs
        {
            get
            {
                return (XTR_EnumImSecurityType)(int)_WorkSiteOject.DefaultVisibility;
            }

            set
            {
                _WorkSiteOject.DefaultVisibility = (IManage.imSecurityType)(int)value;
            }
        }

        public bool Inherited
        {
            get
            {
                return _WorkSiteOject.Inherited;
            }
            set
            {
                _WorkSiteOject.Inherited = value;
            }
        }

        public List<XTR_GroupACL> Groups
        {
            get
            {
                List<XTR_GroupACL> groupAcl = new List<XTR_GroupACL>();

                if (_WorkSiteOject.GroupACLs != null && _WorkSiteOject.GroupACLs.Count > 0)
                {
                    for (int i = 1; i <= _WorkSiteOject.GroupACLs.Count; i++)
                    {
                        XTR_GroupACL item = new XTR_GroupACL();
                        item._WorkSiteOject = _WorkSiteOject.GroupACLs.ItemByIndex(i);
                        groupAcl.Add(item);
                    }
                }

                return groupAcl;
            }

        }

        public void GroupAdd(string name, XTR_EnumAccessRight level)
        {
            _WorkSiteOject.GroupACLs.Add(name, (imAccessRight)(int)level);
        }

        /// <summary>
        /// altera caso ja exista
        /// </summary>
        /// <param name="name"></param>
        /// <param name="level"></param>
        public void UsersAdd(string name, XTR_EnumAccessRight level)
        {
              _WorkSiteOject.UserACLs.Add(name, (imAccessRight)(int)level);   
        }

        /// <summary>
        /// Remove se existir
        /// </summary>
        /// <param name="name"></param>
        public void UsersRemove(string name)
        {
            for (int i = 1; i <= _WorkSiteOject.UserACLs.Count; i++)
            {
               var user = _WorkSiteOject.UserACLs.ItemByIndex(i);
                if(user.User.Name.ToUpper().Equals(name))
                {
                    _WorkSiteOject.UserACLs.RemoveByIndex(i);
                }
            }

        }

        public void GroupsRemove(string name)
        {
            for (int i = 1; i <= _WorkSiteOject.GroupACLs.Count; i++)
            {
                var group = _WorkSiteOject.GroupACLs.ItemByIndex(i);
                if (group.Group.Name.ToUpper().Equals(name))
                {
                    _WorkSiteOject.GroupACLs.RemoveByIndex(i);
                }
            }

        }

        public void UsersClear()
        {
            _WorkSiteOject.UserACLs.Clear();
        }

        public void GroupClear()
        {
            _WorkSiteOject.GroupACLs.Clear();
        }


        public List<XTR_UserACL> Users
        {
            get
            {
                List<XTR_UserACL> usersAcl = new List<XTR_UserACL>();

                if (_WorkSiteOject.UserACLs != null && _WorkSiteOject.UserACLs.Count > 0)
                {
                    for (int i = 1; i <= _WorkSiteOject.UserACLs.Count; i++)
                    {
                        XTR_UserACL item = new XTR_UserACL();
                        item._WorkSiteOject = _WorkSiteOject.UserACLs.ItemByIndex(i);
                        usersAcl.Add(item);
                    }
                }
                return usersAcl;
            }
        }

        internal IManage.IManSecurity _WorkSiteOject { get; set; }
    }
}
