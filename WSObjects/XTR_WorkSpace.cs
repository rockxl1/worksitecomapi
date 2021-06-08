using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XTRWorkSite.Helper;

namespace XTRWorkSite.WSObjects
{
    public class XTR_WorkSpace
    {
        public DateTime CreationDate
        {
            get { return _WorkSiteOject.CreationDate; }
        }

        public DateTime DateModified { get { return _WorkSiteOject.DateModified; } }

        public string Description
        {
            get { return _WorkSiteOject.Description; }
            set { _WorkSiteOject.Description = value; }
        }

        public string EmailPrefix
        {
            get { return _WorkSiteOject.EmailPrefix; }
            set { _WorkSiteOject.EmailPrefix = value; }
        }

        public bool Hidden
        {
            get { return _WorkSiteOject.Hidden; }
            set { _WorkSiteOject.Hidden = value; }
        }

        public string ID
        {
            get { return _WorkSiteOject.ID; }
        }

        public string Name
        {
            get { return _WorkSiteOject.Name; }
            set { _WorkSiteOject.Name = value; }
        }

        public int WorkspaceID
        {
            get { return _WorkSiteOject.WorkspaceID; }
        }

        public int FolderID
        {
            get { return _WorkSiteOject.FolderID; }
        }

        public List<XTR_Folder> Folders
        {
            get
            {
                List<XTR_Folder> folderResult = new List<XTR_Folder>();

                IManage.IManDocumentFolders folders = _WorkSiteOject.DocumentFolders;

                for (int i = 1; i <= folders.Count; i++)
                {
                    XTR_Folder fol = new XTR_Folder();
                    fol._WorkSiteOject = folders.ItemByIndex(i);
                    folderResult.Add(fol);
                }

                return folderResult;
            }
        }

        public XTR_DocumentClass Class
        {
            get
            {
                XTR_DocumentClass classe = new XTR_DocumentClass();
                classe._WorkSiteOject = _WorkSiteOject.Class;
                return classe;
            }
        }

        public XTR_DocumentClass SubClass
        {
            get
            {
                XTR_DocumentClass SubClass = new XTR_DocumentClass();
                SubClass._WorkSiteOject = _WorkSiteOject.SubClass;
                if (SubClass._WorkSiteOject == null)
                    return null;
                return SubClass;
            }
        }

        public XTR_Security Security
        {
            get
            {
                XTR_Security security = new XTR_Security();
                security._WorkSiteOject = _WorkSiteOject.Security;
                return security;
            }
        }

        public XTR_User Owner
        {
            get
            {
                XTR_User own = new XTR_User();
                own._WorkSiteOject = _WorkSiteOject.Owner;
                return own;
            }

            set
            {
                _WorkSiteOject.Owner = value._WorkSiteOject;
            }
        }


        internal IManage.IManWorkspace _WorkSiteOject { get; set; } //a ligação ao worksite está aqui

        public void AddExtendsAttribute(XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID att, object value)
        {
            //nao precisa de remover o attributo 1º como as pastas. o worksapce custom data esta noutra tabela
            _WorkSiteOject.SetAttributeByID(EnumBinding.GetImProfileAttributeID(att), value);
        }

        //igual a anterior com a diferenca que recebe uma lista
        public void AddExtendsAttributes(Dictionary<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID, object> values)
        {
            foreach (var item in values)
            {
                _WorkSiteOject.SetAttributeByID(EnumBinding.GetImProfileAttributeID(item.Key), item.Value);
            }
            
        }

        public object GetExtendsAttribute(XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID att)
        {
            return _WorkSiteOject.GetAttributeValueByID(Helper.EnumBinding.GetImProfileAttributeID(att));
        }

        //public List<XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID>> GetExtendsAttributes()
        //{
        //    return _WorkSiteOject
        //}

        /// <summary>
        /// Save in WorkSite the changes maded
        /// </summary>
        public void Update()
        {
            _WorkSiteOject.Update();
        }

        public void Refresh()
        {
            _WorkSiteOject.Refresh();
        }
    }


}
