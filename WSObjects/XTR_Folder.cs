using IManage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XTRWorkSite.WSObjects.Enums;

namespace XTRWorkSite.WSObjects
{
    public class XTR_Folder
    {
        internal IManage.IManFolder _WorkSiteOject { get; set; }

        public string Description
        {
            get
            {
                return _WorkSiteOject.Description;
            }
            set
            {
                _WorkSiteOject.Description = value;
            }
        }

        public int FolderID
        {
            get {
                return _WorkSiteOject.FolderID; }
        }

        public string Name
        {
            get
            {
                return _WorkSiteOject.Name;
            }
            set
            {
                _WorkSiteOject.Name = value;
            }
        }

        public string EmailPrefix
        {
            get
            {
                return _WorkSiteOject.EmailPrefix;
            }
            set
            {
                _WorkSiteOject.EmailPrefix = value;
            }
        }

        public XTR_WorkSpace Workspace
        {
            get
            {
                XTR_WorkSpace _workspace = new XTR_WorkSpace();
                _workspace._WorkSiteOject = _WorkSiteOject.Workspace;
                return _workspace;
            }
        }

        public XTR_Folder Parent
        {
            get
            {
                XTR_Folder _parent = new XTR_Folder();
                _parent._WorkSiteOject = _WorkSiteOject.Parent;
                return _parent;
            }
        }

        public List<XTR_Folder> SubFolders
        {
            get
            {
                List<XTR_Folder> _subfolders = new List<XTR_Folder>();

                foreach (IManage.IManFolder item in _WorkSiteOject.SubFolders)
                {
                    XTR_Folder cursor = new XTR_Folder();
                    cursor._WorkSiteOject = item;
                    _subfolders.Add(cursor);
                }

                return _subfolders;
            }
        }

        public Enums.XTR_EnumObjectType FolderType
        {
            get
            {
              return (Enums.XTR_EnumObjectType)Enum.Parse(typeof(Enums.XTR_EnumObjectType), _WorkSiteOject.ObjectType.ObjectType.ToString());
            }
        }

        public bool IsSearchFolder
        {
            get
            {
                if (_WorkSiteOject.ObjectType != null && (_WorkSiteOject.ObjectType.ObjectType == IManage.imObjectType.imTypeDocumentSearchFolder || _WorkSiteOject.ObjectType.ObjectType == IManage.imObjectType.imTypeDocumentSearchFolders))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        private List<string> _ParentFoldersNames;
        public List<string> ParentFoldersNames
        {
            get
            {
                if (_ParentFoldersNames == null || _ParentFoldersIds == null)
                {
                    _ParentFoldersIds = new List<int>();
                    _ParentFoldersNames = new List<string>();

                    for (int i = 1; i < _WorkSiteOject.Path.Count; i++)
                    {
                        IManage.IManFolder folder = _WorkSiteOject.Path.ItemByIndex(i);
                        _ParentFoldersNames.Add(folder.Name);
                        _ParentFoldersIds.Add(folder.FolderID);
                    }
                }

                return _ParentFoldersNames;
            }
        }

        private List<int> _ParentFoldersIds;
        public List<int> ParentFoldersIds
        {
            get
            {
                if (_ParentFoldersNames == null || _ParentFoldersIds == null)
                {
                    _ParentFoldersIds = new List<int>();
                    _ParentFoldersNames = new List<string>();

                    for (int i = 1; i < _WorkSiteOject.Path.Count; i++)
                    {
                        IManage.IManFolder folder = _WorkSiteOject.Path.ItemByIndex(i);
                        _ParentFoldersNames.Add(folder.Name);
                        _ParentFoldersIds.Add(folder.FolderID);
                    }
                }

                return _ParentFoldersIds;
            }
        }

        public bool IsEmailFolder
        {
            get
            {
                if (!string.IsNullOrEmpty(_WorkSiteOject.EmailPrefix))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public XTR_Security Security
        {
            get
            {
                XTR_Security _Security = new XTR_Security();
                _Security._WorkSiteOject = _WorkSiteOject.Security;
                return _Security;
            }
        }
      
        public string GetExtendsAttribute(XTRWorkSite.WSObjects.Enums.XTR_EnumNewFolderAttributes att)
        {
            string key = string.Format("iman___{0}", (int)att); 

            for (int i = 1; i <= _WorkSiteOject.AdditionalProperties.Count; i++)
            {

                IManage.IManAdditionalProperty item = _WorkSiteOject.AdditionalProperties.ItemByIndex(i);

                if (key.Equals(item.Name.ToLower()))
                {
                    return item.Value;
                }

            }
            return string.Empty;
        }

        public void SetExtendsAttribute(XTRWorkSite.WSObjects.Enums.XTR_EnumNewFolderAttributes att, string value)
        {
            if(IsSearchFolder)
            {
                IManDocumentSearchFolder searchFolder = (IManDocumentSearchFolder)_WorkSiteOject;
                if(searchFolder.ProfileSearchParameters.Contains((imProfileAttributeID)(int)att))
                {
                    searchFolder.ProfileSearchParameters.ItemByAttribute((imProfileAttributeID)(int)att).Value = value;
                }
                else
                {
                    searchFolder.ProfileSearchParameters.Add((imProfileAttributeID)(int)att, value);
                }
                
              
            }
            else
            {
                string key = string.Format("iMan___{0}", (int)att);

                if (_WorkSiteOject.AdditionalProperties.ContainsByName(key))
                {
                    _WorkSiteOject.AdditionalProperties.RemoveByName(key);
                }

                _WorkSiteOject.AdditionalProperties.Add(key, value);
            }
            
        }

        public void SetExtendsAttributes(Dictionary<XTR_EnumNewFolderAttributes, string> values)
        {
            foreach (var item in values)
            {
                if (IsSearchFolder && !string.IsNullOrEmpty(item.Value)) // 
                {
                    IManDocumentSearchFolder searchFolder = (IManDocumentSearchFolder)_WorkSiteOject;
                    if (searchFolder.ProfileSearchParameters.Contains((imProfileAttributeID)(int)item.Key))
                    {
                        searchFolder.ProfileSearchParameters.ItemByAttribute((imProfileAttributeID)(int)item.Key).Value = item.Value;
                    }
                    else
                    {
                        searchFolder.ProfileSearchParameters.Add((imProfileAttributeID)(int)item.Key, item.Value);
                    }
                }
                else
                {
                    string key = string.Format("iMan___{0}", (int)item.Key);

                    if (_WorkSiteOject.AdditionalProperties.ContainsByName(key))
                    {
                        _WorkSiteOject.AdditionalProperties.RemoveByName(key);
                    }

                    _WorkSiteOject.AdditionalProperties.Add(key, item.Value);
                }
            }
        }

        public List<XTR_Document> GetDocuments()
        {
            List<XTR_Document> result = new List<XTR_Document>();

            for (int i = 1; i <= _WorkSiteOject.Contents.Count; i++)
            {
                XTR_Document doc = new XTR_Document();
                doc._WorkSiteOject = (IManage.IManDocument)_WorkSiteOject.Contents.ItemByIndex(i);
                result.Add(doc);
            }

            return result;
        }
        
        public void Update()
        {
            _WorkSiteOject.Update();
        }
    }
}
