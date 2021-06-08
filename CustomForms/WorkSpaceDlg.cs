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
    public partial class WorkSpaceDlg : Form
    {
        private List<IManWorkspace> _workSpacesGrid = new List<IManWorkspace>();
        private List<IManWorkspace> _workSpaceCache = new List<IManWorkspace>(); //for no index search
        private bool _useIndexer = true;
        private IManSession _sess { get; set; }
        private IManDMS _m_dms { get; set; }
        private IManDatabase _dataBasename { get; set; }
        

        public bool _isCancel = true;
        public IManWorkspace _selectedWorkSpace = null;
        public XTR_WorkSpace _selectedWorkSpaceXTR = null; 

        public WorkSpaceDlg(IManSession session, IManDMS dms, bool useIndexer, IManDatabase dataBase)
        {
            _sess = session;
            _dataBasename = dataBase;
            _useIndexer = useIndexer;
            _m_dms = dms;
            InitializeComponent();
            InitializeDataGrid();

            int total = 214;
            if (!useIndexer)
            {
                total = 9999;
            }

            _sess.MaxRowsForSearch = total;
            _sess.MaxRowsNonSearch = total;

            LoadAllWorkSpaces(total);

            txtDescription.KeyDown += new KeyEventHandler(descp_KeyDown);
        }

        public WorkSpaceDlg(string host, string username, string password, string database, bool trusted, bool useIndexer)
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

            //igual ao outro construtor
            _dataBasename = (IManage.IManDatabase)_sess.Databases.ItemByName(database);
            _sess.Timeout = 30;

            InitializeComponent();
            InitializeDataGrid();

            _useIndexer = useIndexer;
            int total = 214;
            if (!useIndexer)
            {
                total = 9999;
            }

            _sess.MaxRowsForSearch = total;
            _sess.MaxRowsNonSearch = total;

            LoadAllWorkSpaces(total);

            txtDescription.KeyDown += new KeyEventHandler(descp_KeyDown);
        }

        public WorkSpaceDlg(WorkSiteAccess accessXTR, bool useIndexer, string database = null)
        {
            _sess = accessXTR._m_session;
            _dataBasename = accessXTR._m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                _dataBasename = accessXTR._m_dataBase.Where(x => x.Name.ToUpper().Equals(database)).First();
            }

            _m_dms = accessXTR._m_dms ;

            InitializeComponent();
            InitializeDataGrid();

            int total = 214;
            if (!useIndexer)
            {
                total = 9999;
            }

            _useIndexer = useIndexer;
            _sess.MaxRowsForSearch = total;
            _sess.MaxRowsNonSearch = total;

            LoadAllWorkSpaces(total);

            txtDescription.KeyDown += new KeyEventHandler(descp_KeyDown);
        }

        private void descp_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                btnSearch_Click(null, null);
            }
        }

        private void LoadAllWorkSpaces(int total)
        {
            try
            {
                IManage.ManStrings bd = new IManage.ManStrings();
                bd.Add(_dataBasename.Name);

                _sess.MaxRowsForSearch = total;
                _sess.MaxRowsNonSearch = total;

                IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
                IManage.IManWorkspaceSearchParameters param = _m_dms.CreateWorkspaceSearchParameters();

                IManage.IManFolders ws = (IManage.IManFolders)_sess.SearchWorkspaces(bd, proSearch, param);

                _workSpacesGrid = new List<IManWorkspace>();

                for (int i = 1; i <= ws.Count; i++)
                {
                    _workSpacesGrid.Add((IManage.IManWorkspace)ws.ItemByIndex(i));
                }
                _workSpaceCache = _workSpacesGrid;

                InitializeDataGrid();
                LoadValuesInGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error LoadAllWorkSpaces", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeDataGrid()
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.CellDoubleClick += DoubleClickRecord;
            dataGridView1.Columns.Add("ID", "ID");
            dataGridView1.Columns.Add("Name", "Name");
            dataGridView1.Columns.Add("C1", "C1");
            dataGridView1.Columns.Add("C2", "C2");

            dataGridView1.Columns[0].Width = GetWidthFromPercentage(10);
            dataGridView1.Columns[1].Width = GetWidthFromPercentage(60);
            dataGridView1.Columns[2].Width = GetWidthFromPercentage(15);
            dataGridView1.Columns[3].Width = GetWidthFromPercentage(15);
        }

        private int GetWidthFromPercentage(int v)
        {
            int total = dataGridView1.Width;
            return (v * total) / 100;
        }

        private void LoadValuesInGrid()
        {
            foreach (var item in _workSpacesGrid)
            {
                string nomeWorkace = item.Name;

                if(!string.IsNullOrEmpty(item.Description))
                {
                    nomeWorkace = string.Format("{0} <{1}>", nomeWorkace, item.Description);
                }

                 dataGridView1.Rows.Add(item.WorkspaceID, nomeWorkace, item.GetAttributeValueByID(imProfileAttributeID.imProfileCustom1), item.GetAttributeValueByID(imProfileAttributeID.imProfileCustom2));
            }

            lblTotal.Text = string.Format("Total: {0}", _workSpacesGrid.Count());
        }


        private void DoubleClickRecord(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                _selectedWorkSpace = GetSelectedValue();

                if (_selectedWorkSpace != null)
                {
                    _selectedWorkSpaceXTR = new XTR_WorkSpace();
                    _selectedWorkSpaceXTR._WorkSiteOject = _selectedWorkSpace;
                    _isCancel = false;
                    Close();
                }
            
            }
            
        }


        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (_useIndexer)
                {
                    #region Usar o Indexer

                    IManage.ManStrings bd = new IManage.ManStrings();
                    bd.Add(_dataBasename.Name);


                    IManage.IManProfileSearchParameters proSearch = _m_dms.CreateProfileSearchParameters();
                    IManage.IManWorkspaceSearchParameters param = _m_dms.CreateWorkspaceSearchParameters();

                    if(!string.IsNullOrEmpty(txtDescription.Text))
                    {
                        proSearch.AddFullTextSearch(txtDescription.Text, imFullTextSearchLocation.imFullTextAnywhere);
                    }
                    
                    IManage.IManFolders ws = (IManage.IManFolders)_sess.SearchWorkspaces(bd, proSearch, param);

                    _workSpacesGrid = new List<IManWorkspace>();

                    for (int i = 1; i <= ws.Count; i++)
                    {
                        _workSpacesGrid.Add((IManage.IManWorkspace)ws.ItemByIndex(i));
                    }

                    InitializeDataGrid();
                    LoadValuesInGrid();

                    #endregion
                }
                else
                {
                    #region Ir a procura dentro da lista cache

                    _workSpacesGrid = (from cursor in _workSpaceCache
                                       where cursor.Name.ToLower().Contains(txtDescription.Text.ToLower()) ||
                                       cursor.Description.ToLower().Contains(txtDescription.Text.ToLower())
                                       select cursor).ToList();

                    InitializeDataGrid();
                    LoadValuesInGrid();

                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error On Click", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            _selectedWorkSpace = GetSelectedValue();

            if (_selectedWorkSpace == null)
            {
                MessageBox.Show("Please select one WorkSpace", "No WorkSpace selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                _selectedWorkSpaceXTR = new XTR_WorkSpace();
                _selectedWorkSpaceXTR._WorkSiteOject = _selectedWorkSpace;
                _isCancel = false;
                this.Close();
            }
        }

        private IManWorkspace GetSelectedValue()
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                try
                {
                    return _workSpacesGrid.Where(x => x.WorkspaceID.Equals(dataGridView1.SelectedRows[0].Cells[0].Value)).First();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error WS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return null; 
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _isCancel = true;
            this.Close();
        }

        private void txtDescription_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
