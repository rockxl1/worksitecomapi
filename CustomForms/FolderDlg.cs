using IManage;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using XTRWorkSite.WSObjects;

namespace XTRWorkSite.CustomForms
{
    public partial class FolderDlg : Form
    {
        private int _startFolderID;
        private IManSession _sess { get; set; }
        private IManDMS _m_dms { get; set; }
        private IManDatabase _dataBaseName;
        private bool _allowRootSelect;

        public IManFolder _selectedFolder;
        public XTR_Folder _selectFolderXTR;
        public bool _isCancel = true;


        public FolderDlg(int startFolderId, IManSession session, IManDMS dms, IManDatabase dataBasename, bool allowRootSelect)
        {
            InitializeComponent();
            _startFolderID = startFolderId;
            _sess = session;
            _m_dms = dms;
            _dataBaseName =  dataBasename;
            _allowRootSelect = allowRootSelect;

            InitializeTreeView();
            LoadTreeView();
        }

        public FolderDlg(int startFolderId, string host, string username, string password, string database, bool trusted, bool allowRootSelect)
        {
            _m_dms = new IManage.ManDMSClass();
            _sess = _m_dms.Sessions.Add(host);

            if (trusted)
            {
                _sess.TrustedLogin();
            }
            else
            {
                _sess.Login(username, password);
            }

            if (!_sess.Connected)
            {
                throw new Exception("Unable to connect to WorkSite using IManage. Review your credencials and connection.");
            }


            _dataBaseName = (IManage.IManDatabase)_sess.Databases.ItemByName(database);

            //igual ao outro construtor
            _startFolderID = startFolderId;
            _allowRootSelect = allowRootSelect;

            InitializeComponent();
            InitializeTreeView();
            LoadTreeView();

        }

        public FolderDlg(int startFolderId, WorkSiteAccess accessXTR, bool allowRootSelect, string database = null)
        {
            accessXTR.LoadIManageLibrary();

            _allowRootSelect = allowRootSelect;

            _m_dms = accessXTR._m_dms;
            _sess = accessXTR._m_session;

            _dataBaseName = accessXTR._m_dataBase.First();

            if(!string.IsNullOrEmpty(database))
            {
                _dataBaseName = accessXTR._m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            //igual ao outro construtor
            _startFolderID = startFolderId;  
           
            InitializeComponent();
            InitializeTreeView();
            LoadTreeView();
        }

        private void InitializeTreeView()
        {
            treeView1.Nodes.Clear();
            treeView1.MouseDoubleClick += DoubleClickNode;
            treeView1.AfterExpand += LoadNextNodes;
        }

        private void LoadNextNodes(object sender, TreeViewEventArgs e)
        {
            try
            {

                foreach (TreeNode cursor in e.Node.Nodes)
                {
                    // skip if we have already loaded data for this node
                    if (cursor.Tag != null)
                        continue;

                    cursor.Tag = new object();

                    _selectedFolder = _dataBaseName.GetFolder(Convert.ToInt32(cursor.Name));

                    foreach (IManFolder item in _selectedFolder.SubFolders)
                    {
                        cursor.Nodes.Add(item.FolderID.ToString(), item.Name);
                    }

                  
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "LoadSubFolder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DoubleClickNode(object sender, MouseEventArgs e)
        {
            try
            {
                TreeNode node = treeView1.SelectedNode;

                if (!string.IsNullOrEmpty(node.Name))
                {
                    _selectedFolder = _dataBaseName.GetFolder(Convert.ToInt32(node.Name));


                    if (_selectedFolder.Parent == null && !_allowRootSelect) 
                    {
                        _selectedFolder = null;
                        MessageBox.Show("Root folder is forbidden", "Forbidden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        _selectFolderXTR = new XTR_Folder();
                        _selectFolderXTR._WorkSiteOject = _selectedFolder;
                        _isCancel = false;
                        this.Close();
                    }
                    
                }
                else
                {
                    _selectedFolder = null;
                }

               
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void LoadTreeView()
        {
            treeView1.Nodes.Clear();

            IManFolder folder = _dataBaseName.GetFolder(_startFolderID);

            
            TreeNode node = treeView1.Nodes.Add(folder.FolderID.ToString(), folder.Name);
            
            if (folder.SubFolders != null && folder.SubFolders.Count > 0)
            {
                foreach (IManFolder item in folder.SubFolders)
                {
                    node.Nodes.Add(item.FolderID.ToString(), item.Name);
                }
                  
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                TreeNode node = treeView1.SelectedNode;

                if (!string.IsNullOrEmpty(node.Name))
                {
                    _selectedFolder = _dataBaseName.GetFolder(Convert.ToInt32(node.Name));


                    if (_selectedFolder.Parent == null && !_allowRootSelect) 
                    {
                        _selectedFolder = null;
                        MessageBox.Show("Root folder is forbidden", "Forbidden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        _selectFolderXTR = new XTR_Folder();
                        _selectFolderXTR._WorkSiteOject = _selectedFolder;
                        _isCancel = false;
                        this.Close();
                    }
                    
                }
                else
                {
                    _selectedFolder = null;
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _isCancel = true;
            this.Close();
        }
    }
}
