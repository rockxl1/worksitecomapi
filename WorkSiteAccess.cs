using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using IManage;
using XTRWorkSite.WSObjects;
using XTRWorkSite.WSObjects.Enums;
using System.Windows.Forms;


namespace XTRWorkSite
{
    public class WorkSiteAccess : IDisposable
    {
        #region Properties

        //DLL IManage
        public IManage.IManDMS _m_dms;
        public IManage.IManSession _m_session;
        public List<IManage.IManDatabase> _m_dataBase;
        public bool _m_IsConnect = false;

        //DLL IMANADMIN
        public IMANADMIN.NRTDMS _a_dms;
        public IMANADMIN.NRTSession _a_session;
        public List<IMANADMIN.NRTDatabase> _a_dataBase;
        public bool _a_IsConnect = false;


        public string server;
        public string username;
        public string password;
        public bool trustedLogin;
        public List<string> dataBaseName;

        #endregion

        #region Construtions

        public WorkSiteAccess(IManage.IManDMS dms, IManage.IManSession session, IManage.IManDatabase dataBase)
        {
            _m_dms = dms;
            _m_session = session;
            _m_dataBase = new List<IManDatabase>();
            _m_dataBase.Add(dataBase);
            dataBaseName = new List<string>();
            dataBaseName.Add(dataBase.Name);
            _m_IsConnect = true;
        }

        public WorkSiteAccess(IManage.IManDMS dms, IManage.IManSession session, List<IManage.IManDatabase> dataBase)
        {
            _m_dms = dms;
            _m_session = session;
            _m_dataBase = dataBase;
            dataBaseName = new List<string>();

            foreach (var item in dataBase)
            {
                dataBaseName.Add(item.Name);
            }
            _m_IsConnect = true;
        }

        public WorkSiteAccess(string server, string username, string password, string dataBaseName, bool trustedLogin)
        {
            this.server = server;
            this.username = username;
            this.password = password;
            this.dataBaseName = new List<string>();
            this.dataBaseName.Add(dataBaseName);
            this.trustedLogin = trustedLogin;
            LoadIManageLibrary();
        }

        public WorkSiteAccess(string server, string username, string password, List<string> dataBaseName, bool trustedLogin)
        {
            this.server = server;
            this.username = username;
            this.password = password;
            this.dataBaseName = dataBaseName;
            this.trustedLogin = trustedLogin;
            LoadIManageLibrary();
        }

        /// <summary>
        /// AdminLibrary é referente a DLL IMANADMIN que tem metodos que a IManage não tem.
        /// Aqui é carregado a biblioteca, e feita a ligação ao worksite. Foi escolhido esta forma e não no construtor para não
        /// carregar logo a cabeça 2 ligações Iman e IManadmin
        /// Esta DLL irá deixar de ser suportada. Nota que esta no SDK.
        /// </summary>
        protected void LoadAdminLibrary()
        {
            if (!_a_IsConnect)
            {
                _a_dms = new IMANADMIN.NRTDMS();
                _a_session = _a_dms.Sessions.Add(server);

                if (trustedLogin)
                {
                    _a_session.TrustedLogin();
                }
                else
                {
                    _a_session.Login(username, password);
                }

                if (!_a_session.Connected)
                {
                    throw new Exception("Unable to connect to WorkSite using IManAdmin. Review your credencials and connection.");
                }

                _a_IsConnect = true;

                _a_dataBase = new List<IMANADMIN.NRTDatabase>();
                foreach (string item in dataBaseName)
                {
                    _a_dataBase.Add((IMANADMIN.NRTDatabase)_a_session.Databases.Item(item));
                }

            }
        }

        /// <summary>
        /// IManageLibrary é referente a DLL IManage.dll que tem metodos principais.
        /// Aqui é carregado a biblioteca, e feita a ligação ao worksite. Foi escolhido esta forma e não no construtor para não
        /// carregar logo a cabeça 2 ligações Iman e IManadmin
        /// </summary>
        internal void LoadIManageLibrary()
        {
            if (!_m_IsConnect)
            {
               
                _m_dms = new IManage.ManDMSClass();
                _m_session = _m_dms.Sessions.Add(server);

                if (trustedLogin)
                {
                    _m_session.TrustedLogin();
                }
                else
                {
                    _m_session.Login(username, password);
                }

                if (!_m_session.Connected)
                {
                    throw new Exception("Unable to connect to WorkSite using IManage. Review your credencials and connection.");
                }

                username = _m_session.UserID;

                _m_dataBase = new List<IManDatabase>();

                foreach (var item in dataBaseName)
                {
                    _m_dataBase.Add((IManage.IManDatabase)_m_session.Databases.ItemByName(item));
                }


                _m_IsConnect = true;

                _m_session.Timeout = 30;
            }
        }

        public bool CheckConnection()
        {

            LoadAdminLibrary();
            LoadIManageLibrary();
            return true;

        }

        public bool CheckIManageConnection()
        {

            LoadIManageLibrary();
            return true;

        }

        public bool CheckIAdminConnection()
        {
            try
            {
                LoadAdminLibrary();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Methods

        #region Custom Methods

        /// <summary>
        /// Change CustomX Description and Enable state, where customAliasId
        /// </summary>
        /// <param name="customTable">Table Custom ID</param>
        /// <param name="customAliasId">String for CUSTOM_ALIAS key</param>
        /// <param name="customDescription">Description for item</param>
        /// <param name="enable">Enable flag</param>
        /// <param name="parentAliasId">Not require. Parent CUSTOM_ALIAS key. Valid for custom2</param>
        public void ChangeCustomItem(XTRWorkSite.WSObjects.Enums.XTR_EnumAttributeID customTable, string customAliasId, string customDescription, bool isEnable, string parentAliasId, string database = null)
        {
            LoadAdminLibrary();

            IMANADMIN.AttributeID valor = XTRWorkSite.Helper.EnumBinding.GetAttributeID(customTable);

            var bd = _a_dataBase.First();

            
            if (!string.IsNullOrEmpty(database))
            {
                bd = _a_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            if (string.IsNullOrEmpty(parentAliasId))
            {
                bd.ModifyParentCustomField(valor, customAliasId.ToUpper(), customAliasId.ToUpper(), customDescription, isEnable);
            }
            else
            {
                bd.ModifyChildCustomField(valor, parentAliasId.ToUpper(), customAliasId.ToUpper(), parentAliasId.ToUpper(), customAliasId.ToUpper(), customDescription, isEnable);
            }

        }

        /// <summary>
        /// Create new Custom Item in CustomX table.
        /// </summary>
        /// <param name="customTable">Table Custom ID</param>
        /// <param name="customAliasId">String for CUSTOM_ALIAS key</param>
        /// <param name="customDescription">Description for item</param>
        /// <param name="isEnable">Enable flag</param>
        /// <param name="parentAliasId">Not require. Parent CUSTOM_ALIAS key. Valid for custom2</param>
        public void CreateCustomItem(XTRWorkSite.WSObjects.Enums.XTR_EnumAttributeID customTable, string customAliasId, string customDescription, bool isEnable, string parentAliasId, string database = null)
        {
            LoadAdminLibrary();

            IMANADMIN.AttributeID valor = XTRWorkSite.Helper.EnumBinding.GetAttributeID(customTable);

            var bd = _a_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                bd = _a_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            if (string.IsNullOrEmpty(parentAliasId))
            {
                bd.InsertParentCustomField(valor, customAliasId.ToUpper(), customDescription, isEnable);
            }
            else
            {
                bd.InsertChildCustomField(valor, parentAliasId.ToUpper(), customAliasId.ToUpper(), customDescription, isEnable);
            }

        }

        /// <summary>
        /// Delete Custom Item in customTable table
        /// </summary>
        /// <param name="customTable">Table Custom Id</param>
        /// <param name="customAliasId">string for CUSTOM_ALIAS key</param>
        public void DeleteCustomItem(XTRWorkSite.WSObjects.Enums.XTR_EnumAttributeID customTable, string customAliasId, string parentId, string database = null)
        {
            LoadAdminLibrary();
            IMANADMIN.AttributeID valor = XTRWorkSite.Helper.EnumBinding.GetAttributeID(customTable);

            var bd = _a_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                bd = _a_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            if (string.IsNullOrEmpty(parentId))
                bd.DeleteParentCustomField(valor, customAliasId);
            else
                bd.DeleteChildCustomField(valor, parentId, customAliasId);
        }

        /// <summary>
        /// If customTable == Custom1 just one element in list, else customTable == Custom2 one ore more. 
        /// </summary>
        /// <param name="customTable">Table id - XTR_ColumnsId because using new API Imanage</param>
        /// <param name="customAliasId">Custom Alias Key</param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_Custom> GetCustomItem(XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID customTable, string customAliasId, string database = null)
        {
            LoadIManageLibrary();

            List<XTRWorkSite.WSObjects.XTR_Custom> result = new List<XTRWorkSite.WSObjects.XTR_Custom>();

            XTRWorkSite.WSObjects.XTR_Custom customObj = null;

            IManage.imProfileAttributeID valor = XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(customTable);

            var bd = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                bd = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManage.IManCustomAttributes attributes = bd.SearchCustomAttributes(customAliasId, valor, IManage.imSearchAttributeType.imSearchExactMatch, IManage.imSearchEnabledDisabled.imSearchEnabledOrDisabled, true);

            if (attributes.Count != 0)
            {
                for (int i = 1; i <= attributes.Count; i++)
                {
                    //retorna o 1º que encontrar. Só existe um pq estamos a pesquisar pela Key
                    customObj = new WSObjects.XTR_Custom();
                    customObj._WorkSiteOject = attributes.ItemByIndex(i);
                    result.Add(customObj);
                }
            }
            return result;
        }

        public List<XTRWorkSite.WSObjects.XTR_Custom> SearchCustomItem(XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID customTable, string customAliasId, int? size = null, string database = null)
        {
            LoadIManageLibrary();
            SetMaxRowsForSearch(size);

            List<XTRWorkSite.WSObjects.XTR_Custom> result = new List<XTRWorkSite.WSObjects.XTR_Custom>();

            XTRWorkSite.WSObjects.XTR_Custom customObj = null;

            IManage.imProfileAttributeID valor = XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(customTable);

            var bd = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                bd = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManage.IManCustomAttributes attributes = bd.SearchCustomAttributes(customAliasId, valor, IManage.imSearchAttributeType.imSearchBoth, IManage.imSearchEnabledDisabled.imSearchEnabledOrDisabled, false);

            if (attributes.Count != 0)
            {
                for (int i = 1; i <= attributes.Count; i++)
                {
                    //retorna o 1º que encontrar. Só existe um pq estamos a pesquisar pela Key
                    customObj = new WSObjects.XTR_Custom();
                    customObj._WorkSiteOject = attributes.ItemByIndex(i);
                    result.Add(customObj);
                }
            }
            return result;
        }

        #endregion

        #region WorkSpace Methods

        /// <summary>
        /// Get a List of WorkSpaces. Just one parameter is require
        /// </summary>
        /// <param name="parameters">Search parameters</param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_WorkSpace> GetWorkSpaces(XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumWorkSpaceAttributes> parameters, XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> parametersProfile = null, string database = null)
        {
            
            LoadIManageLibrary();

            bool searchDown = true;

            if (parameters != null)
            {
                SetMaxRowsForSearch(parameters.MaxRowCount);
                searchDown = parameters.SearchDown;
            }
            if (parametersProfile != null)
            {
                SetMaxRowsForSearch(parametersProfile.MaxRowCount);
                searchDown = parametersProfile.SearchDown;
            }

            IManage.ManStrings bd = new IManage.ManStrings();

            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }



            List<XTRWorkSite.WSObjects.XTR_WorkSpace> result = null;


            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            IManage.IManWorkspaceSearchParameters param = _m_dms.CreateWorkspaceSearchParameters();


            if (parameters != null && parameters.Parameters != null && parameters.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumWorkSpaceAttributes> cursor in parameters.Parameters)
                {
                    param.Add(XTRWorkSite.Helper.EnumBinding.GetWorkSpaceSearchParameteres(cursor.Column), cursor.Value);
                }
            }

            if (parametersProfile != null && parametersProfile.Parameters != null && parametersProfile.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> cursor in parametersProfile.Parameters)
                {
                    proSearch.Add(XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(cursor.Column), cursor.Value);
                }
            }

            IManage.IManFolders ws = (IManage.IManFolders)_m_session.SearchWorkspaces(bd, proSearch, param);

            result = new List<WSObjects.XTR_WorkSpace>();

            for (int i = 1; i <= ws.Count; i++)
            {
                WSObjects.XTR_WorkSpace xtr_workSpace = new WSObjects.XTR_WorkSpace();
                xtr_workSpace._WorkSiteOject = (IManage.IManWorkspace)ws.ItemByIndex(i);
                result.Add(xtr_workSpace);
            }



            return result;
        }

        /// <summary>
        /// Get a List of WorkSpaces. Just one parameter is require
        /// </summary>
        /// <param name="parameters">Search parameters</param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_WorkSpace> GetWorkSpacesByNameAndFilter(string name, XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumWorkSpaceAttributes> parameters, XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> parametersProfile = null, string database = null)
        {

            LoadIManageLibrary();

            bool searchDown = true;

            if (parameters != null)
            {
                SetMaxRowsForSearch(parameters.MaxRowCount);
                searchDown = parameters.SearchDown;
            }
            if (parametersProfile != null)
            {
                SetMaxRowsForSearch(parametersProfile.MaxRowCount);
                searchDown = parametersProfile.SearchDown;
            }

            IManage.ManStrings bd = new IManage.ManStrings();

            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }



            List<XTRWorkSite.WSObjects.XTR_WorkSpace> result = null;


            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            IManage.IManWorkspaceSearchParameters param = _m_dms.CreateWorkspaceSearchParameters();

            proSearch.AddFullTextSearch(name, imFullTextSearchLocation.imFullTextAnywhere);

            if (parameters != null && parameters.Parameters != null && parameters.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumWorkSpaceAttributes> cursor in parameters.Parameters)
                {
                    param.Add(XTRWorkSite.Helper.EnumBinding.GetWorkSpaceSearchParameteres(cursor.Column), cursor.Value);
                }
            }

            if (parametersProfile != null && parametersProfile.Parameters != null && parametersProfile.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> cursor in parametersProfile.Parameters)
                {
                    proSearch.Add(XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(cursor.Column), cursor.Value);
                }
            }

            IManage.IManFolders ws = (IManage.IManFolders)_m_session.SearchWorkspaces(bd, proSearch, param);

            result = new List<WSObjects.XTR_WorkSpace>();

            for (int i = 1; i <= ws.Count; i++)
            {
                WSObjects.XTR_WorkSpace xtr_workSpace = new WSObjects.XTR_WorkSpace();
                xtr_workSpace._WorkSiteOject = (IManage.IManWorkspace)ws.ItemByIndex(i);
                result.Add(xtr_workSpace);
            }



            return result;
        }

        /// <summary>
        /// Get a List of WorkSpaces. Just one parameter is require
        /// </summary>
        /// <param name="parameters">Search parameters</param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_WorkSpace> GetWorkSpacesByName(string name, string database = null)
        {
            LoadIManageLibrary();

            bool searchDown = true;

            IManage.ManStrings bd = new IManage.ManStrings();

            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }


            List<XTRWorkSite.WSObjects.XTR_WorkSpace> result = null;


            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            IManage.IManWorkspaceSearchParameters param = _m_dms.CreateWorkspaceSearchParameters();

            proSearch.AddFullTextSearch(name, imFullTextSearchLocation.imFullTextAnywhere);

            IManage.IManFolders ws = (IManage.IManFolders)_m_session.SearchWorkspaces(bd, proSearch, param);

            result = new List<WSObjects.XTR_WorkSpace>();

            for (int i = 1; i <= ws.Count; i++)
            {
                WSObjects.XTR_WorkSpace xtr_workSpace = new WSObjects.XTR_WorkSpace();
                xtr_workSpace._WorkSiteOject = (IManage.IManWorkspace)ws.ItemByIndex(i);
                result.Add(xtr_workSpace);
            }

            return result;
        }

        public void AddWorkSpaceToRecentWorkSpaces(XTR_WorkSpace workspace)
        {
            LoadIManageLibrary();
            _m_session.WorkArea.RecentWorkspaces.AddWorkspace(workspace._WorkSiteOject);

        }

        /// <summary>
        /// Preformance of this method is bad
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public XTR_WorkSpace GetWorkSpace(string id, string database = null)
        {
            LoadIManageLibrary();

            XTR_WorkSpace result = null;

            IManage.ManStrings bd = new IManage.ManStrings();


            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }

            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            IManage.IManWorkspaceSearchParameters param = _m_dms.CreateWorkspaceSearchParameters();

            param.Add(imFolderAttributeID.imFolderID, id);

            IManage.IManFolders ws = (IManage.IManFolders)_m_session.SearchWorkspaces(bd, proSearch, param);

            if (ws != null && ws.Count > 0)
            {
                result = new XTR_WorkSpace();
                result._WorkSiteOject = (IManWorkspace)ws.ItemByIndex(1);
            }

            return result;
        }

        public List<XTRWorkSite.WSObjects.XTR_WorkSpace> GetMyRecentWorkSpaces(int top)
        {
            LoadIManageLibrary();
            SetMaxRowsForSearch(top);

            List<XTRWorkSite.WSObjects.XTR_WorkSpace> result = new List<WSObjects.XTR_WorkSpace>();

            IManage.IManWorkspaces wks = _m_session.WorkArea.RecentWorkspaces;

            for (int i = 1; i <= wks.Count; i++)
            {
                WSObjects.XTR_WorkSpace xtr_workSpace = new WSObjects.XTR_WorkSpace();
                xtr_workSpace._WorkSiteOject = (IManage.IManWorkspace)wks.ItemByIndex(i);
                result.Add(xtr_workSpace);

                if (i >= top)
                    break;
            }

            SetMaxRowsForSearch(9999);

            return result;

        }

        //public XTR_WorkSpace CreateWorkSpace(XTR_Security securityXML, XTR_Security securityTXT, XTR_WorkSpace ws, List<XTR_Folder> folders, string database = null)
        //{
        //    LoadIManageLibrary();

        //    var db = _m_dataBase.First();
        //    if (!string.IsNullOrEmpty(database))
        //    {
        //        db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
        //    }

        //    IManage.ManStrings bd = new IManage.ManStrings();


        //    if (!string.IsNullOrEmpty(database))
        //    {
        //        bd.Add(database);
        //    }
        //    else
        //    {
        //        foreach (var item in dataBaseName)
        //        {
        //            bd.Add(item);
        //        }
        //    }

        //    IManWorkspace workspace = null;
       
        //    try
        //    {
        //        workspace = db.CreateWorkspace();
                
        //        workspace.Name = ws.Name;
        //        workspace.Description = ws.Description;
        //        workspace.Hidden = false;
        //        workspace.Owner = ws.Owner._WorkSiteOject;

        //        foreach (var item in  ws.GetExtendsAttributes())
        //        {
        //            workspace.SetAttributeByID(Helper.EnumBinding.GetImProfileAttributeID(item.Column), item.Value);
        //        }


        //        if (securityTXT != null) //ou um ou outro
        //        {
        //            #region Set Security on WORKSPACE from TXT

        //            workspace.Security.DefaultVisibility = (imSecurityType)(int)securityTXT.DefaultOrSharedAs;

        //            if (securityTXT.Users.Count > 0)
        //            {
        //                foreach (var item in securityTXT.Users)
        //                {
        //                    if (!workspace.Security.UserACLs.Contains(item.User.Name))
        //                    {
        //                        workspace.Security.UserACLs.Add(item.User.Name, (imAccessRight)item.Right);
        //                    }
        //                }
        //            }

        //            if (securityTXT.Groups.Count > 0)
        //            {
        //                foreach (var item in securityTXT.Groups)
        //                {
        //                    if (!workspace.Security.GroupACLs.Contains(item.Group.Name))
        //                    {
        //                        workspace.Security.GroupACLs.Add(item.Group.Name, (imAccessRight)item.Right);
        //                    }
        //                }
        //            }

        //            #endregion
        //        }

        //        if (securityXML != null)
        //        {
        //            #region Set Security on WORKSPACE from XML

        //            workspace.Security.DefaultVisibility = (imSecurityType)(int)securityXML.DefaultOrSharedAs;

        //            if (securityXML.Users.Count > 0)
        //            {
        //                foreach (var item in securityXML.Users)
        //                {
        //                    if (!workspace.Security.UserACLs.Contains(item.User.Name))
        //                    {
        //                        workspace.Security.UserACLs.Add(item.User.Name, (imAccessRight)item.Right);
        //                    }
        //                }
        //            }

        //            if (securityXML.Groups.Count > 0)
        //            {
        //                foreach (var item in securityXML.Groups)
        //                {
        //                    if (!workspace.Security.GroupACLs.Contains(item.Group.Name))
        //                    {
        //                        workspace.Security.GroupACLs.Add(item.Group.Name, (imAccessRight)item.Right);
        //                    }
        //                }
        //            }

        //            #endregion
        //        }

        //        workspace.Update(); //commit ja.

        //        if (folders.Count > 0)
        //        {
        //            CreateFoldersForWorkSpace(workspace, folders, bd);

        //            #region Create Folders
                   

        //            //foreach (var item in folders)
        //            //{
        //            //    IManage.IManDocumentFolders sourceFolders = (IManage.IManDocumentFolders)workspace.SubFolders;
        //            //    IManage.IManDocumentSearchFolders sourceSearchFolders = (IManage.IManDocumentSearchFolders)workspace.SubFolders;

        //            //    IManage.IManDocumentFolder m_folder = null;
        //            //    IManage.IManDocumentSearchFolder m_search_folder = null;

        //            //    if (!item.IsSearchFolder)
        //            //    {
        //            //        #region Document Folder

        //            //        if (item.Security.Inherited)
        //            //        {
        //            //            m_folder = sourceFolders.AddNewDocumentFolderInheriting(item.Name, item.Description);
        //            //        }
        //            //        else
        //            //        {
        //            //            m_folder = sourceFolders.AddNewDocumentFolder(item.Name, item.Description);
        //            //        }

        //            //        foreach (var cursor in item.Security.Users)
        //            //        {
        //            //            if (!m_folder.Security.UserACLs.Contains(cursor.User.Name))
        //            //            {
        //            //                m_folder.Security.UserACLs.Add(cursor.User.Name, (imAccessRight)cursor.Right);
        //            //            }
        //            //        }

        //            //        foreach (var cursor in item.Security.Groups)
        //            //        {
        //            //            if (!m_folder.Security.GroupACLs.Contains(cursor.Group.Name))
        //            //            {
        //            //                m_folder.Security.GroupACLs.Add(cursor.Group.Name, (imAccessRight)cursor.Right);
        //            //            }
        //            //        }

        //            //        m_folder.Update();
        //            //        m_folder.Refresh();

        //            //        #endregion
        //            //    }
        //            //    else
        //            //    {
        //            //        #region Search Folder

        //            //        IManProfileSearchParameters objProfileSearchParams = _m_dms.CreateProfileSearchParameters();

        //            //        if (item.SearchFolderAttributes != null)
        //            //        {
        //            //            foreach (var searchAtt in item.SearchFolderAttributes.Parameters)
        //            //            {
        //            //                objProfileSearchParams.Add((imProfileAttributeID)(int)searchAtt.Column, searchAtt.Value);
        //            //            }
        //            //        }

        //            //        if (item.Security.Inherited)
        //            //        {
        //            //            m_search_folder = sourceSearchFolders.AddNewDocumentSearchFolderInheriting(item.Name, item.Description, bd, objProfileSearchParams);
        //            //        }
        //            //        else
        //            //        {
        //            //            m_search_folder = sourceSearchFolders.AddNewDocumentSearchFolder(item.Name, item.Description, bd, objProfileSearchParams);
        //            //        }


        //            //        foreach (var cursor in item.Security.Users)
        //            //        {
        //            //            if (!m_search_folder.Security.UserACLs.Contains(cursor.User.Name))
        //            //            {
        //            //                m_search_folder.Security.UserACLs.Add(cursor.User.Name, (imAccessRight)cursor.Right);
        //            //            }
        //            //        }

        //            //        foreach (var cursor in item.Security.Groups)
        //            //        {
        //            //            if (!m_search_folder.Security.GroupACLs.Contains(cursor.Group.Name))
        //            //            {
        //            //                m_search_folder.Security.GroupACLs.Add(cursor.Group.Name, (imAccessRight)cursor.Right);
        //            //            }
        //            //        }

        //            //        m_search_folder.Update();
        //            //        m_search_folder.Refresh();

        //            //        #endregion
        //            //    }
        //            //}

        //            #endregion
        //        }

        //        XTR_WorkSpace work = new XTR_WorkSpace();
        //        work._WorkSiteOject = workspace;
        //        return work;
        //    }
        //    catch (Exception ex)
        //    {
        //        //rollback
        //        if (workspace != null)
        //        {
        //            db.DeleteWorkspace(workspace.WorkspaceID);
        //        }
        //        throw ex;
        //    }
        //}

        //private void CreateFoldersForWorkSpace(IManFolder startUp, List<XTR_Folder> folders, IManage.ManStrings bd)
        //{
        //    IManage.IManDocumentFolders sourceFolders = (IManage.IManDocumentFolders)startUp.SubFolders;
        //    IManage.IManDocumentSearchFolders sourceSearchFolders = (IManage.IManDocumentSearchFolders)startUp.SubFolders;

        //    IManage.IManDocumentFolder m_folder = null;
        //    IManage.IManDocumentSearchFolder m_search_folder = null;

        //    foreach (XTR_Folder item in folders)
        //    {

        //        if (!item.IsSearchFolder)
        //        {
        //            #region Document Folder

        //            if (item.Security.Inherited)
        //            {
        //                m_folder = sourceFolders.AddNewDocumentFolderInheriting(item.Name, item.Description);
        //                if(item.IsEmailFolder)
        //                {
        //                    m_folder.EmailPrefix = string.Format("FID{0}", m_folder.FolderID);
        //                }
        //            }
        //            else
        //            {
        //                m_folder = sourceFolders.AddNewDocumentFolder(item.Name, item.Description);
        //                if (item.IsEmailFolder)
        //                {
        //                    m_folder.EmailPrefix = string.Format("FID{0}", m_folder.FolderID);
        //                }
        //            }

        //            foreach (var cursor in item.Security.Users)
        //            {
        //                if (!m_folder.Security.UserACLs.Contains(cursor.User.Name))
        //                {
        //                    m_folder.Security.UserACLs.Add(cursor.User.Name, (imAccessRight)cursor.Right);
        //                }
        //            }

        //            foreach (var cursor in item.Security.Groups)
        //            {
        //                if (!m_folder.Security.GroupACLs.Contains(cursor.Group.Name))
        //                {
        //                    m_folder.Security.GroupACLs.Add(cursor.Group.Name, (imAccessRight)cursor.Right);
        //                }
        //            }

        //            m_folder.Update();
        //            m_folder.Refresh();

        //            #endregion
        //        }
        //        else
        //        {
        //            #region Search Folder

        //            IManProfileSearchParameters objProfileSearchParams = _m_dms.CreateProfileSearchParameters();

        //            if (item.SearchFolderAttributes != null)
        //            {
        //                foreach (var searchAtt in item.SearchFolderAttributes.Parameters)
        //                {
        //                    objProfileSearchParams.Add((imProfileAttributeID)(int)searchAtt.Column, searchAtt.Value);
        //                }
        //            }

        //            if (item.Security.Inherited)
        //            {
        //                m_search_folder = sourceSearchFolders.AddNewDocumentSearchFolderInheriting(item.Name, item.Description, bd, objProfileSearchParams);
        //            }
        //            else
        //            {
        //                m_search_folder = sourceSearchFolders.AddNewDocumentSearchFolder(item.Name, item.Description, bd, objProfileSearchParams);
        //            }


        //            foreach (var cursor in item.Security.Users)
        //            {
        //                if (!m_search_folder.Security.UserACLs.Contains(cursor.User.Name))
        //                {
        //                    m_search_folder.Security.UserACLs.Add(cursor.User.Name, (imAccessRight)cursor.Right);
        //                }
        //            }

        //            foreach (var cursor in item.Security.Groups)
        //            {
        //                if (!m_search_folder.Security.GroupACLs.Contains(cursor.Group.Name))
        //                {
        //                    m_search_folder.Security.GroupACLs.Add(cursor.Group.Name, (imAccessRight)cursor.Right);
        //                }
        //            }

        //            m_search_folder.Update();
        //            m_search_folder.Refresh();

        //            #endregion
        //        }

        //        if (item.SubFolders != null && item.SubFolders.Count > 0)
        //        {
        //            if (m_folder != null)
        //            {
        //                CreateFoldersForWorkSpace(m_folder, item.SubFolders, bd);
        //            }
        //            else if (m_search_folder != null)
        //            {
        //                CreateFoldersForWorkSpace(m_search_folder, item.SubFolders, bd);
        //            }

        //        }
        //    }

        //}

        #endregion

        #region Folder Methods

        /// <summary>
        /// Get Folders by Query. If query empty, return all
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_Folder> GetFolders(XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumFolderAttributeId> parameters, string database = null)
        {

            List<XTRWorkSite.WSObjects.XTR_Folder> result = null;

            LoadIManageLibrary();

            SetMaxRowsForSearch(parameters.MaxRowCount);

            IManage.ManStrings bd = new IManage.ManStrings();

            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }

            IManage.IManFolderSearchParameters proSearch = _m_dms.CreateFolderSearchParameters();

            if (parameters != null && parameters.Parameters != null && parameters.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumFolderAttributeId> cursor in parameters.Parameters)
                {
                    proSearch.Add(XTRWorkSite.Helper.EnumBinding.GetImFolderAttributeID(cursor.Column), cursor.Value);
                }
            }


            IManage.IManFolders folders = _m_session.WorkArea.SearchFolders(bd, proSearch);

            result = new List<WSObjects.XTR_Folder>();

            for (int i = 1; i <= folders.Count; i++)
            {
                WSObjects.XTR_Folder xtr_workSpace = new WSObjects.XTR_Folder();
                xtr_workSpace._WorkSiteOject = folders.ItemByIndex(i);
                result.Add(xtr_workSpace);
            }

            return result;
        }

        /// <summary>
        /// Get Folder by ID
        /// </summary>
        /// <param name="idFolder"></param>
        /// <returns></returns>
        public XTRWorkSite.WSObjects.XTR_Folder GetFolder(int idFolder, string database = null)
        {
            LoadIManageLibrary();

            XTRWorkSite.WSObjects.XTR_Folder folder = new WSObjects.XTR_Folder();

            var db = _m_dataBase.First();
            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            folder._WorkSiteOject = db.GetFolder(idFolder); 

            return folder;
        }

        /// <summary>
        /// The folder is created under WorkSpace
        /// </summary>
        /// <param name="workspace">Folder under WorkSpace</param>
        /// <param name="folderName"></param>
        /// <param name="folderDescription"></param>
        /// <param name="Inheriting"></param>
        /// <param name="additionalProperties"></param>
        public XTRWorkSite.WSObjects.XTR_Folder CreateFolder(XTRWorkSite.WSObjects.XTR_WorkSpace workspace, string folderName, string folderDescription, bool Inheriting, Dictionary<XTRWorkSite.WSObjects.Enums.XTR_EnumNewFolderAttributes, string> additionalProperties)
        {
            LoadIManageLibrary();

            IManage.IManDocumentFolders source = workspace._WorkSiteOject.DocumentFolders;

            IManage.IManDocumentFolder m_folder = null;

            if (Inheriting)
            {
                m_folder = source.AddNewDocumentFolderInheriting(folderName, folderDescription);
            }
            else
            {
                m_folder = source.AddNewDocumentFolder(folderName, folderDescription);
            }

            for (int i = 25; i <= 54; i++) //pura martelada visto que a api nao tem .Attributes quando obj = workspace
            {
                imProfileAttributeID customCursor = (imProfileAttributeID)i; //

                object value = workspace._WorkSiteOject.GetAttributeValueByID(customCursor);

                object temValor = workspace._WorkSiteOject.GetAttributeByID(customCursor); //para caso seja bool, vem null

                if (value != null && !string.IsNullOrEmpty(value.ToString()) && temValor != null)
                {

                    m_folder.AdditionalProperties.Add("iMan___" + (int)i, value.ToString());
                }
            }

            foreach (var item in additionalProperties)
            {
                m_folder.AdditionalProperties.Add("iMan___" + (int)item.Key, item.Value);  //visto a API aceitar string em vez de enumerado  
            }

            if (m_folder.AdditionalProperties.Count > 0)
            {
                m_folder.Update();
            }

            XTRWorkSite.WSObjects.XTR_Folder result = new XTRWorkSite.WSObjects.XTR_Folder();
            result._WorkSiteOject = m_folder;

            return result;
        }

        /// <summary>
        /// Create Folder under folder
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="folderName"></param>
        /// <param name="folderDescription"></param>
        /// <param name="Inheriting"></param>
        /// <param name="additionalProperties"></param>
        /// <returns></returns>
        //[Obsolete]
        //public int CreateFolder(XTRWorkSite.WSObjects.XTR_Folder folder, string folderName, string folderDescription, bool Inheriting, Dictionary<XTRWorkSite.WSObjects.Enums.XTR_EnumNewFolderAttributes, string> additionalProperties)
        //{
        //    LoadIManageLibrary();

        //    IManage.IManDocumentFolders source = (IManage.IManDocumentFolders)folder._WorkSiteOject.SubFolders;

        //    IManage.IManDocumentFolder m_folder = null;

        //    if (Inheriting)
        //    {
        //        m_folder = source.AddNewDocumentFolderInheriting(folderName, folderDescription);

        //        if (folder._WorkSiteOject.CustomProperties != null)
        //        {
        //            for (int i = 1; i <= folder._WorkSiteOject.AdditionalProperties.Count; i++)
        //            {
        //                IManAdditionalProperty cusmP = folder._WorkSiteOject.AdditionalProperties.ItemByIndex(i);

        //                if (!m_folder.AdditionalProperties.ContainsByName(cusmP.Name))
        //                {
        //                    m_folder.AdditionalProperties.Add(cusmP.Name, cusmP.Value);
        //                }

        //            }
        //            m_folder.Update();
        //        }
        //    }
        //    else
        //    {
        //        m_folder = source.AddNewDocumentFolder(folderName, folderDescription);
        //    }



        //    foreach (var item in additionalProperties)
        //    {
        //        m_folder.AdditionalProperties.Add("iMan___" + (int)item.Key, item.Value);  //visto a API aceitar string em vez de enumerado  
        //    }
        //    if (additionalProperties.Count() > 0)
        //    {
        //        m_folder.Update();
        //    }

        //    return m_folder.FolderID;
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="folderName"></param>
        /// <param name="folderDescription"></param>
        /// <param name="Inheriting"></param>
        /// <param name="additionalProperties"></param>
        /// <param name="returnAllData">retorna no object XTR_Folder as subpastas, workspace etc. + rápido</param>
        /// <returns></returns>
        public XTRWorkSite.WSObjects.XTR_Folder CreateFolder(XTRWorkSite.WSObjects.XTR_Folder folder, string folderName, string folderDescription, bool Inheriting, Dictionary<XTRWorkSite.WSObjects.Enums.XTR_EnumNewFolderAttributes, string> additionalProperties)
        {
            LoadIManageLibrary();

            if (folder._WorkSiteOject is IManWorkspace)
            {
                #region if is worksapce

                XTR_WorkSpace ws = GetWorkSpace(folder.FolderID.ToString());

                return CreateFolder(ws, folderName, folderDescription, Inheriting, additionalProperties);

                #endregion
            }
            else
            {
                #region if is folder


                IManDocumentFolders source = (IManage.IManDocumentFolders)folder._WorkSiteOject.SubFolders;

                IManage.IManDocumentFolder m_folder = null;

                if (Inheriting)
                {
                    m_folder = source.AddNewDocumentFolderInheriting(folderName, folderDescription);

                    if (folder._WorkSiteOject.CustomProperties != null)
                    {
                        for (int i = 1; i <= folder._WorkSiteOject.AdditionalProperties.Count; i++)
                        {
                            IManAdditionalProperty cusmP = folder._WorkSiteOject.AdditionalProperties.ItemByIndex(i);

                            if (!m_folder.AdditionalProperties.ContainsByName(cusmP.Name))
                            {
                                m_folder.AdditionalProperties.Add(cusmP.Name, cusmP.Value);
                            }

                        }
                        m_folder.Update();
                    }
                }
                else
                {
                    m_folder = source.AddNewDocumentFolder(folderName, folderDescription);
                }


                foreach (var item in additionalProperties)
                {
                    m_folder.AdditionalProperties.Add("iMan___" + (int)item.Key, item.Value);  //visto a API aceitar string em vez de enumerado  
                }
                if (additionalProperties.Count() > 0)
                {
                    m_folder.Update();
                }

                XTR_Folder folderX = new XTR_Folder();
                folderX._WorkSiteOject = m_folder;

                return folderX;

                #endregion
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="souceFoFolderID">Id da pasta a mover</param>
        /// <param name="targetFolderId">Id do futuro pai da pasta a mover</param>
        public void MoveFolder(int souceFoFolderID, int targetFolderId, string database = null)
        {
            LoadIManageLibrary();

            IManDatabase db = _m_dataBase.First();
            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManFolder folderSource = db.GetFolder(souceFoFolderID);
            IManFolder folderTarget = db.GetFolder(targetFolderId);

            folderSource.MoveTo(folderTarget.SubFolders);
        }

        public void DeleteFolder(int id, string database = null)
        {
            
            LoadIManageLibrary();
            IManDatabase db = _m_dataBase.First();
            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManFolder folder = db.GetFolder(id);
            IManFolder parent = folder.Parent;

            if(parent != null)
            {
                parent.SubFolders.RemoveByObject(folder);
                parent.Update();
            }   
        }

        #endregion

        #region Documents Methods

        /// <summary>
        /// Return one or a List of Documents.
        /// </summary>
        /// <param name="parameters">Search parameters</param>
        /// <returns>Document List</returns>
        public List<XTRWorkSite.WSObjects.XTR_Document> GetDocuments(XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> parameters, XTRWorkSite.WSObjects.Enums.XTR_EnumImSearchEmail searchType, string database = null)
        {
            LoadIManageLibrary();

            SetMaxRowsForSearch(parameters.MaxRowCount);

            IManage.ManStrings bd = new IManage.ManStrings();

            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }

            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();

            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            proSearch.SearchEmail = (imSearchEmail)searchType;

            if (parameters != null && parameters.Parameters != null && parameters.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> cursor in parameters.Parameters)
                {
                    proSearch.Add(XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(cursor.Column), cursor.Value);
                }
            }

            IManage.IManDocuments documentos = (IManage.IManDocuments)_m_session.SearchDocuments(bd, proSearch, true);

            

            for (int i = 1; i <= documentos.Count; i++)
            {
                WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                xtr_document._WorkSiteOject = ((IManage.IManDocument)documentos.ItemByIndex(i));
                result.Add(xtr_document);
            }

            return result;
        }

        /// <summary>
        /// Return one or a List of Documents.
        /// </summary>
        /// <param name="parameters">Search parameters</param>
        /// <returns>Document List</returns>
        public List<XTRWorkSite.WSObjects.XTR_Document> GetDocumentsSort(XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> parameters, XTRWorkSite.WSObjects.Enums.XTR_EnumImSearchEmail searchType, string database = null)
        {
            LoadIManageLibrary();

            SetMaxRowsForSearch(parameters.MaxRowCount);

            IManage.ManStrings bd = new IManage.ManStrings();

            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }

            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();

            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            proSearch.SearchEmail = (imSearchEmail)searchType;
            

            if (parameters != null && parameters.Parameters != null && parameters.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> cursor in parameters.Parameters)
                {
                    proSearch.Add(XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(cursor.Column), cursor.Value);
                }
            }


            XTR_Sort objSorter = new XTR_Sort();

            IManage.IManDocuments documentos = (IManage.IManDocuments)_m_session.SearchDocuments(bd, proSearch, true);

            documentos.Sort(objSorter);


            for (int i = 1; i <= documentos.Count; i++)
            {
                WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                xtr_document._WorkSiteOject = ((IManage.IManDocument)documentos.ItemByIndex(i));
                result.Add(xtr_document);
            }

            return result;
        }


        /// <summary>
        /// nao é possivel pesquisar por folder ID
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="searchType"></param>
        /// <param name="database"></param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_Document> GetDocumentsAdvance(XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> parameters, XTRWorkSite.WSObjects.Enums.XTR_EnumImSearchEmail searchType, string database = null)
        {
            LoadIManageLibrary();

            SetMaxRowsForSearch(parameters.MaxRowCount);

            IManage.ManStrings bd = new IManage.ManStrings();

            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }


            IManAndQuery andQuery = new ManAndQueryClass();

            int? anterior = null;

            Dictionary<int,IManFieldQuery> andQueryField = new Dictionary<int, IManFieldQuery>();
            foreach (var item in parameters.Parameters.OrderBy(x=>x.Column))
            {
                int token = (int)item.Column;

                if (!anterior.HasValue || anterior.Value != (int)item.Column)
                {
                    andQueryField[token] = andQuery.AddFieldQuery((imProfileAttributeID)(int)item.Column); //ands
                    andQueryField[token].Values.Add(item.Value);
                }
                else
                {
                    andQueryField[token].Values.Add(item.Value);
                }

                anterior = (int)item.Column;
            }

            string strResults = string.Empty;
            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();

            // Pass in the ManAndQuery parameter to run the search
            IManage.IManDocuments documentos = (IManDocuments)_m_session.WorkArea.SearchDocumentsEx(bd, andQuery);

            for (int i = 1; i <= documentos.Count; i++)
            {
                WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                xtr_document._WorkSiteOject = ((IManage.IManDocument)documentos.ItemByIndex(i));
                result.Add(xtr_document);
            }


            return result;
        }

        public List<XTRWorkSite.WSObjects.XTR_Document> GetDocumentsByString(string searchString, XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> parameters, XTRWorkSite.WSObjects.Enums.XTR_EnumImSearchEmail searchType, string database = null)
        {
            LoadIManageLibrary();

            SetMaxRowsForSearch(parameters.MaxRowCount);

            IManage.ManStrings bd = new IManage.ManStrings();


            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }

            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();

            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            proSearch.AddFullTextSearch(searchString, imFullTextSearchLocation.imFullTextAnywhere);


            proSearch.SearchEmail = (imSearchEmail)searchType;


            if (parameters != null && parameters.Parameters != null && parameters.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> cursor in parameters.Parameters)
                {
                    proSearch.Add(XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(cursor.Column), cursor.Value);
                }
            }



            IManage.IManDocuments documentos = (IManage.IManDocuments)_m_session.SearchDocuments(bd, proSearch, true);

            for (int i = 1; i <= documentos.Count; i++)
            {
                WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                xtr_document._WorkSiteOject = ((IManage.IManDocument)documentos.ItemByIndex(i));
                result.Add(xtr_document);
            }

            return result;
        }

        public List<XTRWorkSite.WSObjects.XTR_Document> GetDocumentsByString(string searchString, XTRWorkSite.WSObjects.XTR_SearchParameters<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> parameters, XTRWorkSite.WSObjects.Enums.XTR_EnumImSearchEmail searchType, XTRWorkSite.WSObjects.Enums.XTR_EnumFullTextSearchLocation searchLocation, string database = null)
        {
            LoadIManageLibrary();

            SetMaxRowsForSearch(parameters.MaxRowCount);

            IManage.ManStrings bd = new IManage.ManStrings();


            if (!string.IsNullOrEmpty(database))
            {
                bd.Add(database);
            }
            else
            {
                foreach (var item in dataBaseName)
                {
                    bd.Add(item);
                }
            }

            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();

            IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
            proSearch.AddFullTextSearch(searchString, (imFullTextSearchLocation)searchLocation);


            proSearch.SearchEmail = (imSearchEmail)searchType;


            if (parameters != null && parameters.Parameters != null && parameters.Parameters.Count > 0)
            {
                foreach (XTRWorkSite.WSObjects.XTR_Parameter<XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID> cursor in parameters.Parameters)
                {
                    proSearch.Add(XTRWorkSite.Helper.EnumBinding.GetImProfileAttributeID(cursor.Column), cursor.Value);
                }
            }



            IManage.IManDocuments documentos = (IManage.IManDocuments)_m_session.SearchDocuments(bd, proSearch, true);

            for (int i = 1; i <= documentos.Count; i++)
            {
                WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                xtr_document._WorkSiteOject = ((IManage.IManDocument)documentos.ItemByIndex(i));
                result.Add(xtr_document);
            }

            return result;
        }

        /// <summary>
        /// Return Documents inside a folder
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="maxRouCount"></param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_Document> GetDocuments(XTRWorkSite.WSObjects.XTR_Folder folder, int? maxRouCount)
        {
            LoadIManageLibrary();

            SetMaxRowsForSearch(maxRouCount);
            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();

            for (int i = 1; i <= folder._WorkSiteOject.Contents.Count; i++)
            {
                WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                xtr_document._WorkSiteOject = ((IManage.IManDocument)folder._WorkSiteOject.Contents.ItemByIndex(i));
                result.Add(xtr_document);
            }

            SetMaxRowsForSearch(9999);

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public List<XTRWorkSite.WSObjects.XTR_Document> GetLastDocumentsWorkList(int size)
        {
            LoadIManageLibrary();
            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();
            SetMaxRowsForSearch(size);

            IManage.IManContents conteudo = _m_session.WorkArea.Worklist;
            
            for (int i = 1; i <= conteudo.Count; i++)
            {
                IManage.IManContent item = conteudo.ItemByIndex(i);

                if (item is IManage.IManDocument)
                {
                    WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                    xtr_document._WorkSiteOject = ((IManage.IManDocument)item);
                    result.Add(xtr_document);

                    if (size <= result.Count())
                        break;
                }
            }

            SetMaxRowsForSearch(9999);

            return result;
        }

        /// <summary>
        /// Atenção isto irá colocar o ficheiro no path.
        /// </summary>
        /// <param name="numer"></param>
        /// <param name="vers"></param>
        /// <param name="path"></param>
        public void CheckOutDocument(int numer, int vers, string path, string message = null, string database = null)
        {
            LoadIManageLibrary();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManage.IManDocument doc = db.GetDocument(numer, vers);

            if (doc.IsOperationAllowed(imDocumentOperation.imCheckOutDocumentOp))
            {
                doc.CheckOut(path, IManage.imCheckOutOptions.imReplaceExistingFile, DateTime.Now, message);
            }
            else
            {
                throw new Exception("Operação CheckOut não é válida para este documento");
            }


        }

        public object CheckInDocument(int numer, int vers, string path, XTRWorkSite.WSObjects.Enums.XTR_CheckInTypes checkInOPT, string database = null)
        {
            LoadIManageLibrary();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManage.IManDocument doc = db.GetDocument(numer, vers);

            if (doc.IsOperationAllowed(imDocumentOperation.imCheckInDocumentOp))
            {
                object error = new object();
                doc.CheckIn(path, XTRWorkSite.Helper.EnumBinding.GetCheckInDisposition(checkInOPT), imCheckinOptions.imDontKeepCheckedOut, ref error);
                return error;
            }
            else
            {
                throw new Exception("Operação CheckIn não é válida");
            }

        }

        /// <summary>
        /// if isNativeFormat = false then is in HTML
        /// </summary>
        /// <param name="numer"></param>
        /// <param name="vers"></param>
        /// <param name="path"></param>
        /// <param name="isNativeFormat"></param>
        public void OpenDocument(int numer, int vers, string path, bool isNativeFormat, string database = null)
        {
            LoadIManageLibrary();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManage.IManDocument doc = db.GetDocument(numer, vers);

            if (isNativeFormat)
                doc.GetCopy(path, imGetCopyOptions.imNativeFormat);
            else
                doc.GetCopy(path, imGetCopyOptions.imHTMLFormat);
        }

        public XTR_Document CreateNewDocument(XTRWorkSite.WSObjects.XTR_Folder pasta, string ownerUsername, string type, string fileName,
            string classe, string path, XTRWorkSite.WSObjects.Enums.XTR_CheckInTypes checkType, XTRWorkSite.WSObjects.Enums.XTR_CheckInOption checkInOpt, Dictionary<XTR_EnumProfileAttributeID, object> extendsAtt = null, bool autoDetectType = false)
        {
            LoadIManageLibrary();

            IManDocument oDoc = ((IManDocuments2)pasta._WorkSiteOject.Contents).CreateDocument();

            oDoc.SetAttributeByID(IManage.imProfileAttributeID.imProfileAuthor, ownerUsername);
            oDoc.SetAttributeByID(IManage.imProfileAttributeID.imProfileOperator, ownerUsername);

            if (autoDetectType)
            {
                IManDocumentType docType = _m_dataBase.First().GetDocumentTypeFromPath(path);
                if (docType != null)
                {
                    type = docType.Name;
                }
            }

            oDoc.SetAttributeByID(IManage.imProfileAttributeID.imProfileType, type);
            oDoc.SetAttributeByID(IManage.imProfileAttributeID.imProfileClass, classe);


            for (int i = 1; i <= pasta._WorkSiteOject.AdditionalProperties.Count; i++)
            {
                try
                {
                    IManAdditionalProperty prop = pasta._WorkSiteOject.AdditionalProperties.ItemByIndex(i);
                    if (prop != null)
                    {
                        int numProp = Convert.ToInt32(prop.Name.ToLower().Replace("iman___", "")); 
                        oDoc.SetAttributeByID((imProfileAttributeID)numProp, prop.Value);
                    }
                }
                catch (Exception exx)
                {

                }

            }

            if (extendsAtt != null)
            {
                foreach (var item in extendsAtt)
                {
                    oDoc.SetAttributeByID(Helper.EnumBinding.GetImProfileAttributeID(item.Key), item.Value);
                }
            }

            oDoc.Description = fileName;

            object CheckInResults = new object();

            oDoc.CheckIn(path, Helper.EnumBinding.GetCheckInDisposition(checkType), Helper.EnumBinding.GetCheckinOpt(checkInOpt), CheckInResults);
            oDoc.Update();

            try
            {

                IManProfile profile = oDoc.Profile;

                //set do security
                oDoc.Security.DefaultVisibility = pasta._WorkSiteOject.Security.DefaultVisibility;

                oDoc.Security.GroupACLs.Clear();
                for (int i = 1; i <= pasta._WorkSiteOject.Security.GroupACLs.Count; i++)
                {
                    IManGroupACL acl = pasta._WorkSiteOject.Security.GroupACLs.ItemByIndex(i);
                    oDoc.Security.GroupACLs.Add(acl.Group.Name, acl.Right);
                }

                oDoc.Security.UserACLs.Clear();
                for (int i = 1; i <= pasta._WorkSiteOject.Security.UserACLs.Count; i++)
                {
                    IManUserACL acl = pasta._WorkSiteOject.Security.UserACLs.ItemByIndex(i);
                    oDoc.Security.UserACLs.Add(acl.User.Name, acl.Right);
                }

                oDoc.Update();

            }
            catch (Exception ex)
            {
                string message = ex.ToString();
            }


            XTR_Document doc = new XTR_Document();
            doc._WorkSiteOject = oDoc;

            return doc;
        }

        public XTRWorkSite.WSObjects.XTR_Document GetDocument(int number, int vers, string database = null)
        {
            LoadIManageLibrary();

            WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }
            
            xtr_document._WorkSiteOject = ((IManage.IManDocument)db.GetDocument(number, vers));
            
            return xtr_document;
        }


        public List<XTRWorkSite.WSObjects.XTR_Document> GetCheckOutDocuments(int size)
        {
            LoadIManageLibrary();
            List<XTRWorkSite.WSObjects.XTR_Document> result = new List<WSObjects.XTR_Document>();

            SetMaxRowsForSearch(size);

            IManage.IManContents conteudo = _m_session.WorkArea.CheckedOutList;

            for (int i = 1; i <= conteudo.Count; i++)
            {
                IManage.IManContent item = conteudo.ItemByIndex(i);

                if (item is IManage.IManDocument)
                {
                    WSObjects.XTR_Document xtr_document = new WSObjects.XTR_Document();
                    xtr_document._WorkSiteOject = ((IManage.IManDocument)item);
                    result.Add(xtr_document);

                    if (size <= result.Count())
                        break;
                }
            }

            SetMaxRowsForSearch(9999);

            return result;

        }


        /// <summary>
        /// Move o documento
        /// </summary>
        /// <param name="number"></param>
        /// <param name="version"></param>
        /// <param name="folderTargetId"></param>
        /// <param name="inheritPremissions"></param>
        /// <param name="removeFromSource"></param>
        /// <param name="database"></param>
        public void MoveDocument(int number, int version, int folderTargetId, bool inheritPremissions, bool removeFromSource, string database = null, bool? isDebug = null)
        {
            LoadIManageLibrary();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManDocument oDoc = db.GetDocument(number, version);


            IManFolder folderTarget = db.GetFolder(folderTargetId);

            IManDocuments targetDocuments = (IManDocuments)folderTarget.Contents;

            targetDocuments.AddDocumentReference(oDoc);

            if (inheritPremissions)
            {
                #region Refile 

                oDoc.Security.DefaultVisibility = folderTarget.Security.DefaultVisibility;

                oDoc.Security.GroupACLs.Clear();
                for (int i = 1; i <= folderTarget.Security.GroupACLs.Count; i++)
                {
                    IManGroupACL acl = folderTarget.Security.GroupACLs.ItemByIndex(i);
                    oDoc.Security.GroupACLs.Add(acl.Group.Name, acl.Right);
                }

                oDoc.Security.UserACLs.Clear();
                for (int i = 1; i <= folderTarget.Security.UserACLs.Count; i++)
                {
                    IManUserACL acl = folderTarget.Security.UserACLs.ItemByIndex(i);
                    oDoc.Security.UserACLs.Add(acl.User.Name, acl.Right);
                }

                #endregion
            }

            oDoc.Update();
            oDoc.Refresh();


            if (removeFromSource)
            {
                #region 1º remover da origem

                if (isDebug.HasValue && isDebug.Value)
                {
                    MessageBox.Show("Total de pastas: " + oDoc.Folders.Count);
                }

                for (int x = 1; x <= oDoc.Folders.Count; x++)
                {
                    IManFolder fol = oDoc.Folders.ItemByIndex(x);
                    fol.Refresh();

                    if (isDebug.HasValue && isDebug.Value)
                    {
                        MessageBox.Show("Pasta: " + fol.Name);
                    }

                    if (fol.FolderID != folderTargetId)
                    {
                        if (isDebug.HasValue && isDebug.Value)
                        {
                            MessageBox.Show("Pasta do qual tem o ficheiro para apagar");
                        }

                        if (fol.Contents != null && fol.Contents.Count > 0)
                        {
                            if (isDebug.HasValue && isDebug.Value)
                            {
                                MessageBox.Show("Vou remover");
                            }

                            //if (fol.Contents.Contains(oDoc))
                            //{
                            fol.Contents.RemoveByObject(oDoc);
                            fol.Update();

                            if (isDebug.HasValue && isDebug.Value)
                            {
                                MessageBox.Show("Removi");
                            }

                            //}
                        }
                    }
                }
                oDoc.Update();
                oDoc.Refresh();
                #endregion
            }
        }

        /// <summary>
        /// Move o documento
        /// </summary>
        /// <param name="number"></param>
        /// <param name="version"></param>
        /// <param name="folderTargetId"></param>
        /// <param name="inheritPremissions"></param>
        /// <param name="removeFromSource"></param>
        /// <param name="database"></param>
        public List<Tuple<int, int, string>> MoveDocuments(List<Tuple<int,int>> numberVersion, int folderTargetId, bool inheritPremissions, bool removeFromSource, string database = null, bool? isDebug = null)
        {
            List<Tuple<int, int, string>> erros = new List<Tuple<int, int, string>>();

            LoadIManageLibrary();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManFolder folderTarget = db.GetFolder(folderTargetId);

            IManDocuments targetDocuments = (IManDocuments)folderTarget.Contents;

            foreach (var item in numberVersion)
            {
                int number = item.Item1;
                int version = item.Item2;

                try
                {
                 

                    IManDocument oDoc = db.GetDocument(number, version);

                    targetDocuments.AddDocumentReference(oDoc);

                    if (inheritPremissions)
                    {
                        #region Refile 

                        oDoc.Security.DefaultVisibility = folderTarget.Security.DefaultVisibility;

                        oDoc.Security.GroupACLs.Clear();
                        for (int i = 1; i <= folderTarget.Security.GroupACLs.Count; i++)
                        {
                            IManGroupACL acl = folderTarget.Security.GroupACLs.ItemByIndex(i);
                            oDoc.Security.GroupACLs.Add(acl.Group.Name, acl.Right);
                        }

                        oDoc.Security.UserACLs.Clear();
                        for (int i = 1; i <= folderTarget.Security.UserACLs.Count; i++)
                        {
                            IManUserACL acl = folderTarget.Security.UserACLs.ItemByIndex(i);
                            oDoc.Security.UserACLs.Add(acl.User.Name, acl.Right);
                        }

                        #endregion
                    }

                    oDoc.Update();
                    oDoc.Refresh();


                    if (removeFromSource)
                    {
                        #region 1º remover da origem

                        if (isDebug.HasValue && isDebug.Value)
                        {
                            MessageBox.Show("Total de pastas: " + oDoc.Folders.Count);
                        }

                        for (int x = 1; x <= oDoc.Folders.Count; x++)
                        {
                            IManFolder fol = oDoc.Folders.ItemByIndex(x);
                            fol.Refresh();

                            if (isDebug.HasValue && isDebug.Value)
                            {
                                MessageBox.Show("Pasta: " + fol.Name);
                            }

                            if (fol.FolderID != folderTargetId)
                            {
                                if (isDebug.HasValue && isDebug.Value)
                                {
                                    MessageBox.Show("Pasta do qual tem o ficheiro para apagar");
                                }

                                if (fol.Contents != null && fol.Contents.Count > 0)
                                {
                                    if (isDebug.HasValue && isDebug.Value)
                                    {
                                        MessageBox.Show("Vou remover");
                                    }

                                    //if (fol.Contents.Contains(oDoc))
                                    //{
                                    fol.Contents.RemoveByObject(oDoc);
                                    fol.Update();

                                    if (isDebug.HasValue && isDebug.Value)
                                    {
                                        MessageBox.Show("Removi");
                                    }

                                    //}
                                }
                            }
                        }
                        oDoc.Update();
                        oDoc.Refresh();
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    erros.Add(new Tuple<int, int, string>(number, version, ex.Message));
                }
               
            }

            return erros;
        }

        /// <summary>
        /// Move o documento
        /// </summary>
        /// <param name="number"></param>
        /// <param name="version"></param>
        /// <param name="folderTargetId"></param>
        /// <param name="inheritPremissions"></param>
        /// <param name="removeFromSource"></param>
        /// <param name="database"></param>
        public void RemoveFromFolder(int number, int version, int folderTargetId, string database = null, bool? isDebug = null)
        {
            LoadIManageLibrary();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManDocument oDoc = db.GetDocument(number, version);

            IManFolder folderTarget = db.GetFolder(folderTargetId);

            IManDocuments targetDocuments = (IManDocuments)folderTarget.Contents;

            targetDocuments.RemoveByObject(oDoc);

            oDoc.Update();
            oDoc.Refresh();

        }

        /// <summary>
        /// Apaga um documento no Worksite
        /// </summary>
        /// <param name="number"></param>
        /// <param name="version"></param>
        public void DeleteDocument(int number, int version)
        {
            LoadIManageLibrary();

            var db = _m_dataBase.First();

            db.DeleteDocument(number, version);
        }

        #endregion

        #region Users Methods

        public XTRWorkSite.WSObjects.XTR_User GetUserName(string username, string database = null)
        {
            LoadIManageLibrary();
            XTRWorkSite.WSObjects.XTR_User user = new XTRWorkSite.WSObjects.XTR_User();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            user._WorkSiteOject = (db.GetUser(username));
            return user;
        }

        public List<XTRWorkSite.WSObjects.XTR_User> GetUsers(string query, string database = null)
        {
            LoadIManageLibrary();
            List<XTRWorkSite.WSObjects.XTR_User> users = new List<XTRWorkSite.WSObjects.XTR_User>();

            var db = _m_dataBase.First();
            
            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }
            db.Session.MaxRowsForSearch = 9999;
            db.Session.MaxRowsNonSearch = 9999;
            IManUsers result = db.SearchUsers(query, imSearchAttributeType.imSearchBoth, true);

            for (int i = 1; i <= result.Count; i++)
            {
                XTRWorkSite.WSObjects.XTR_User user = new WSObjects.XTR_User();
                user._WorkSiteOject = (result.ItemByIndex(i));
                users.Add(user);
            }

            return users;
        }

        public void ChangePassword(string oldPassword, string newPassword)
        {
            LoadIManageLibrary();
            _m_session.ChangePassword(oldPassword, newPassword);
        }

        public void CreateUser(string username, string name, string password)
        {
            LoadAdminLibrary();

            var bd = _a_dataBase.First();

            var user = bd.CreateUser(username);

            user.FullName = name;
            user.Password = password;
            user.PasswordExpires = true;
            
            user.Update();
        }

        public void SetUserCustomProperty(string username, string custom1)
        {
            LoadAdminLibrary();

            var bd = _a_dataBase.First();
            var user = bd.GetUser(username);
            user.Custom1 = custom1;
            user.Update();
        }

        #endregion

        #region Generic Methods

        public List<XTR_DocumentType> GetWorksiteTypes(string searchString, string database = null)
        {
            LoadIManageLibrary();
            List<XTR_DocumentType> types = new List<XTR_DocumentType>();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManDocumentTypes docTypes = db.SearchDocumentTypes(searchString, imSearchAttributeType.imSearchBoth, true);

            for (int i = 1; i <= docTypes.Count; i++)
            {
                XTR_DocumentType xType = new XTR_DocumentType();
                xType._WorkSiteOject = docTypes.Item(i);
                types.Add(xType);
            }

            return types;
        }

        public List<XTR_DocumentClass> GetWorksiteClasses(string searchString, string database = null)
        {
            LoadIManageLibrary();
            List<XTR_DocumentClass> result = new List<XTR_DocumentClass>();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManDocumentClasses docClass = db.SearchDocumentClasses(searchString, imSearchAttributeType.imSearchBoth, true);

            for (int i = 1; i <= docClass.Count; i++)
            {
                IManDocumentClass item = docClass.ItemByIndex(i);
                XTR_DocumentClass docCl = new XTR_DocumentClass();
                docCl._WorkSiteOject = (item);
                result.Add(docCl);
            }


            return result;
        }

        public List<XTR_Group> GetGroups(string name, string database = null)
        {
            List<XTR_Group> result = new List<XTR_Group>();

            LoadIManageLibrary();

            var db = _m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                db = _m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            IManGroups groups = db.SearchGroups(name, imSearchAttributeType.imSearchBoth, true);

            for (int i = 1; i <= groups.Count; i++)
            {
                IManGroup item = groups.ItemByIndex(i);
                XTR_Group grp = new XTR_Group();
                grp._WorkSiteOject = (item);
                result.Add(grp);
            }

            return result;

        }

        public void CreateHistory(XTRWorkSite.WSObjects.XTR_Document document, XTRWorkSite.WSObjects.Enums.XTR_EnumHistEvent histEvent, int duration, int pagesPrinted, string application, string comment, string location, string customString1, string customString2, object customNumber1, object customNumber2, object customNumber3)
        {
            LoadIManageLibrary();
            document._WorkSiteOject.HistoryList.Add(XTRWorkSite.Helper.EnumBinding.GetHistEvent(histEvent), duration, pagesPrinted, application, comment, location, customString1, customString2, customNumber1, customNumber2, customNumber3);
            document._WorkSiteOject.Update();
        }

        public XTR_DocumentType DetectExtensionAuto(string pathFile)
        {
            LoadIManageLibrary();

            XTR_DocumentType result = null;

            try
            {
                IManDocumentType docType = _m_dataBase.First().GetDocumentTypeFromPath(pathFile);
                if (docType != null)
                {
                    result = new XTR_DocumentType();
                    result._WorkSiteOject = (docType);
                }  
            }
            catch { }

            return result;
        }

        /// <summary>
        /// Max value = 9999 (WorkSite limitation)
        /// </summary>
        /// <param name="num"> gre 0 and les 9999 </param>
        public void SetMaxRowsForSearch(int? num)
        {
            if (num.HasValue)
            {
                if (_m_session != null)
                {
                    _m_session.MaxRowsForSearch = num.Value;
                    _m_session.MaxRowsNonSearch = num.Value;
                }

                if (_a_session != null)
                {
                    _a_session.MaxRowsForSearch = num.Value;
                }
            }

        }

        public void SetAllVersions(bool allVersions)
        {
            if (_m_session != null)
            {
                _m_session.AllVersions = allVersions;
            }
        }

        /// <summary>
        /// Close WorkSite Connections
        /// </summary>
        /// 
        public void Dispose()
        {
            if (_a_IsConnect)
            {
                _a_dms.CloseApplication();
                _a_IsConnect = false;
            }


            if (_m_IsConnect)
            {
                _m_dms.Close();
                _m_IsConnect = false;
            }
        }


        #endregion

        #endregion
    }
}
