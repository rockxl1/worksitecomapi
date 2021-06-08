using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_Document
    {
        internal IManage.IManDocument _WorkSiteOject { get; set; }

        public DateTime AccessTime
        {
            get
            {
                return _WorkSiteOject.AccessTime;
            }
        }

        public bool CheckedOut
        {
            get
            {
                return _WorkSiteOject.CheckedOut;
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

        public string Comment
        {
            get { return _WorkSiteOject.Comment; }
            set { _WorkSiteOject.Comment = value; }
        }

        public DateTime CreationDate
        {
            get {return _WorkSiteOject.CreationDate;}
        }

        public string Description {
            get { return _WorkSiteOject.Description; }
            set { _WorkSiteOject.Description = value; }
        }

        public DateTime EditDate {
            get { return _WorkSiteOject.EditDate; }
        }

        public string Extension {
            get { return _WorkSiteOject.Extension; }
        }

        public bool Locked {
            get { return _WorkSiteOject.Locked; }
        }

        public bool DeclaredAsRecord {
            get { return (bool)_WorkSiteOject.GetAttributeByID(IManage.imProfileAttributeID.imProfileFrozen); }
        }

        public string Name {
            get { return _WorkSiteOject.Name; }
        }

        public int Number {
            get { return _WorkSiteOject.Number; }
        }

        public int Size {
            get { return _WorkSiteOject.Size; }
        }

        public XTR_DocumentClass SubClass
        {
            get
            {
                XTR_DocumentClass subClass = new XTR_DocumentClass();
                subClass._WorkSiteOject = _WorkSiteOject.SubClass;
                return subClass;
            }
        }

        public XTR_DocumentType Type
        {
            get
            {
                XTR_DocumentType type = new XTR_DocumentType();
                type._WorkSiteOject = _WorkSiteOject.Type;
                return type;
            }
        }

        public int Version
        {
            get
            {
                return _WorkSiteOject.Version;
            }
        }

        public XTR_User Author
        {
            get
            {
                XTR_User autor = new XTR_User();
                autor._WorkSiteOject = _WorkSiteOject.Author;
                return autor;
            }
        }

        public XTR_User LastUser
        {
            get
            {
                XTR_User luser = new XTR_User();
                luser._WorkSiteOject = _WorkSiteOject.LastUser;
                return luser;
            }
        }

        public XTR_User InUseBy
        {
            get
            {
                XTR_User InUseBy = new XTR_User();
                InUseBy._WorkSiteOject = _WorkSiteOject.InUseBy;
                return InUseBy;
            }
        }

        public string CheckoutLocation
        {
            get
            {
                return _WorkSiteOject.CheckoutLocation;
            }
        }


        public List<XTR_Document> OtherVersions()
        {
            List<XTR_Document> otherVersions = new List<XTR_Document>();

            for (int i = 1; i <= _WorkSiteOject.Versions.Count; i++)
            {
                XTR_Document record = new XTR_Document();
                record._WorkSiteOject = (IManage.IManDocument)_WorkSiteOject.Versions.ItemByIndex(i);
                otherVersions.Add(record);
            }

            return otherVersions;
        }

        public List<XTR_Document> Related()
        {
            List<XTR_Document> related = new List<XTR_Document>();

            for (int i = 1; i <= _WorkSiteOject.RelatedDocuments.Count; i++)
            {
                XTR_Document record = new XTR_Document();
                record._WorkSiteOject = (IManage.IManDocument)_WorkSiteOject.RelatedDocuments.ItemByIndex(i);
                related.Add(record);
            }

            return related;
        }

        public List<XTR_DocumentHistory> GetHistory()
        {
            List<XTR_DocumentHistory> result = new List<XTR_DocumentHistory>();

            for (int i = 1; i <= _WorkSiteOject.HistoryList.Count; i++)
            {
                XTR_DocumentHistory docHistor = new XTR_DocumentHistory();
                docHistor._WorkSiteOject = _WorkSiteOject.HistoryList.ItemByIndex(i);
                result.Add(docHistor);
            }

            return result;
        }


        public void AddExtendsAttribute(XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID att, object value)
        {
            _WorkSiteOject.SetAttributeByID(Helper.EnumBinding.GetImProfileAttributeID(att), value);
        }

        public void AddExtendsAttributes(Dictionary<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID, object> values)
        {
            foreach (var item in values)
            {
                _WorkSiteOject.SetAttributeByID(Helper.EnumBinding.GetImProfileAttributeID(item.Key), item.Value);
            }
        }

        public object GetExtendsAttribute(XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID att)
        {
            return _WorkSiteOject.GetAttributeValueByID(Helper.EnumBinding.GetImProfileAttributeID(att));
        }

        public List<XTR_Folder> GetFolders()
        {
            List<XTR_Folder> result = new List<XTR_Folder>();

            try
            {
                if (_WorkSiteOject.Folders != null && _WorkSiteOject.Folders.Count > 0)
                {
                    for (int i = 1; i <= _WorkSiteOject.Folders.Count; i++)
                    {
                        XTR_Folder fol = new XTR_Folder();
                        fol._WorkSiteOject = _WorkSiteOject.Folders.ItemByIndex(i);
                        result.Add(fol);
                    }
                }
            }
            catch (Exception)
            {

            }

            return result;
        }

        public XTR_Security Security
        {
            get
            {
                XTR_Security sec = new XTR_Security();
                sec._WorkSiteOject = _WorkSiteOject.Security;
                return sec;
            }
        }
      
        public void Update()
        {
            _WorkSiteOject.Update();
        }

        public void UnlockContent()
        {
            _WorkSiteOject.UnlockContent();
        }

        public void Refresh()
        {
            _WorkSiteOject.Refresh();
        }

    }
}
