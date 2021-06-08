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
    public partial class UserDlg : Form
    {
        public bool _isCancel = true;
        private List<XTR_User> users;
        public XTR_User _selectuserXTR;
        public IManUser _selectuser;

        private IManDMS _m_dms { get; set; }
        private IManSession _sess { get; set; }
        private IManDatabase _dataBasename { get; set; }
        
        private List<IManCustomAttribute> _searchResult { get; set; }

        public UserDlg(IManSession session, IManDMS dms, IManDatabase dataBase)
        {
            InitializeComponent();
            _sess = session;
            _m_dms = dms;
            _dataBasename = dataBase;
           
            InitializeGrid();
        }

        public UserDlg(string host, string username, string password, string database, bool trusted)
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
            InitializeGrid();
        }

        public UserDlg(WorkSiteAccess accessXTR, string database = null)
        {
            _sess = accessXTR._m_session;
            _dataBasename = accessXTR._m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                _dataBasename = accessXTR._m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            _m_dms = accessXTR._m_dms;
            InitializeComponent();
            InitializeGrid();
        }

        private void InitializeGrid()
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.CellDoubleClick += Mouse_doubleClick;
            dataGridView1.Columns.Add("UserId", "User Id");
            dataGridView1.Columns.Add("FullName", "Full Name");
            
            _sess.MaxRowsForSearch = 9999;
            _sess.MaxRowsNonSearch = 9999;

            LoadValues(string.Empty);

            txtSearchUser.KeyDown += new KeyEventHandler(descp_KeyDown);
        }

        public void LoadValues(string query)
        {
            try
            {
                dataGridView1.Rows.Clear();

                if (users == null || users.Count == 0)
                {
                    users = new List<XTR_User>();

                   var cursor = _dataBasename.SearchUsers(string.Empty, imSearchAttributeType.imSearchBoth, true);

                   for (int i = 1; i <= cursor.Count; i++)
                   {
                       var aux = new XTR_User();
                       aux._WorkSiteOject=(cursor.ItemByIndex(i));
                       users.Add(aux);
                   }
                }

                // users = _workSiteAccess.GetUsers(query);
                List<XTR_User> result = users;

                if (!string.IsNullOrEmpty(query))
                    result = result.Where(x => x.FullName.ToUpper().Contains(query.ToUpper()) ||
                        x.Name.ToUpper().Contains(query.ToUpper()) ||
                        x.Email.ToUpper().Contains(query.ToUpper())).ToList();

                lblItemsCount.Text = result.Count().ToString();

                foreach (var item in result)
                {
                    dataGridView1.Rows.Add(item.Name, item.FullName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Load users", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Mouse_doubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    try
                    {
                        _selectuserXTR = users.Where(x => x.Name.Equals(dataGridView1.SelectedRows[0].Cells[0].Value)).First();
                        _selectuser = _selectuserXTR._WorkSiteOject;
                        _isCancel = false;
                        Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error. Description: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
            
        }

        private void descp_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                LoadValues(txtSearchUser.Text);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadValues(txtSearchUser.Text);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _isCancel = true;
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                try
                {
                    _selectuserXTR = users.Where(x => x.Name.Equals(dataGridView1.SelectedRows[0].Cells[0].Value)).First();
                    _selectuser = _selectuserXTR._WorkSiteOject;
                    _isCancel = false;
                    Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error. Description: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void txtSearchUser_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
