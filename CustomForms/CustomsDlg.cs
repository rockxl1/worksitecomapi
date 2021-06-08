using IManage;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using XTRWorkSite.Helper;
using XTRWorkSite.WSObjects;
using XTRWorkSite.WSObjects.Enums;

namespace XTRWorkSite.CustomForms
{
    public partial class CustomsDlg : Form
    {
        public bool _isCancel = true;
        public IManCustomAttribute _selectCustom;
        public XTR_Custom _selectCustomXTR;

        private IManDMS _m_dms { get; set; }
        private IManSession _sess { get; set; }
        private IManDatabase _dataBasename { get; set; }
        private imProfileAttributeID _CustomX { get; set; }

        private List<IManCustomAttribute> _searchResult { get; set; }

        public bool AddEnable = false;
        public IMANADMIN.NRTDatabase _dataBasenameAdmin { get; set; }


        //private List<IManCustomAttribute> customXActuaList { get; set; }
        //private List<IManCustomAttribute> customXChillActualList { get; set; }
        private string _parentValue;

        public CustomsDlg(imProfileAttributeID custom, IManSession session, IManDMS dms, IManDatabase dataBase, string parentValue = null, bool? enableAdd = null)
        {
            _CustomX = custom;
            _parentValue = parentValue;
            _sess = session;
            _dataBasename = dataBase;
            _m_dms = dms;
            InitializeComponent();
            InitializeGrid();

            if (enableAdd.HasValue && enableAdd.Value)
            {
                btnAdd.Enabled = true;
            }
            else
            {
                btnAdd.Enabled = false;
            }
        }

        public CustomsDlg(int custom, string host, string username, string password, string database, bool trusted, string parentValue = null, bool? enableAdd = null)
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

            _CustomX = (imProfileAttributeID)custom;
            _parentValue = parentValue;
            InitializeComponent();

            InitializeGrid();

            if (enableAdd.HasValue && enableAdd.Value)
            {
                btnAdd.Enabled = true;
            }
            else
            {
                btnAdd.Enabled = false;
            }
        }

        public CustomsDlg(XTR_EnumProfileAttributeID custom, WorkSiteAccess accessXTR, string parentValue = null, string database = null, bool? enableAdd = null)
        {
            _CustomX = EnumBinding.GetImProfileAttributeID(custom);
            _parentValue = parentValue;
            _sess = accessXTR._m_session;
            _dataBasename = accessXTR._m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                _dataBasename = accessXTR._m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            _m_dms = accessXTR._m_dms;
            InitializeComponent();

            InitializeGrid();

            if (enableAdd.HasValue && enableAdd.Value)
            {
                btnAdd.Enabled = true;
            }
            else
            {
                btnAdd.Enabled = false;
            }
        }

        private void InitializeGrid()
        {
            int total = 214;

            _sess.MaxRowsForSearch = total;
            _sess.MaxRowsNonSearch = total;

            dataGridViewCustom.Columns.Clear();
            dataGridViewCustom.Rows.Clear();
            dataGridViewCustom.Columns.Add("Alias", "Alias");
            dataGridViewCustom.Columns.Add("Description", "Description");
            dataGridViewCustom.Columns[0].Width = 120;
            dataGridViewCustom.CellDoubleClick += DoubleClickRecord;
            LoadValues(string.Empty);

            txtSearch.KeyDown += new KeyEventHandler(descp_KeyDown);
        }

        public void EnableAddBtn(bool enableBtn)
        {
            btnAdd.Enabled = enableBtn;
        }

        private void descp_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                btnFindNow_Click_1(null, null);
            }
        }

        public void LoadValues(string query)
        {
            try
            {
                dataGridViewCustom.Rows.Clear();

                List<IManCustomAttribute> result = GetValues(query);
                _searchResult = result;

                foreach (var item in result)
                {
                    dataGridViewCustom.Rows.Add(item.Name, item.Description);

                }
                lblItemsCount.Text = result.Count().ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private List<IManCustomAttribute> GetValues(string query)
        {
            List<IManCustomAttribute> result = new List<IManCustomAttribute>();

            try
            {

                if (_CustomX == imProfileAttributeID.imProfileCustom2 && !string.IsNullOrEmpty(_parentValue))
                {
                    #region Pesquisa por C2, verificar se é C2 dentro do C1 ou liver


                    //procurar dentro do proprio
                    IManCustomAttributes c1 = _dataBasename.SearchCustomAttributes(_parentValue, imProfileAttributeID.imProfileCustom1, imSearchAttributeType.imSearchExactMatch, imSearchEnabledDisabled.imSearchEnabledOnly, false);

                    if (c1.Count != 2)
                    {
                        IManCustomAttribute c1Att = c1.ItemByIndex(1);
                        IManCustomAttributes c2 = c1Att.GetChildList(query, imSearchAttributeType.imSearchBoth, imSearchEnabledDisabled.imSearchEnabledOnly, false);

                        for (int i = 1; i <= c2.Count; i++)
                        {
                            result.Add(c2.ItemByIndex(i));
                        }

                    }
                    else
                    {
                        MessageBox.Show(string.Format("{0} not found or found more than 1 by api", _parentValue), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }



                    #endregion
                }
                else
                {
                    IManCustomAttributes cX = _dataBasename.SearchCustomAttributes(query, _CustomX, imSearchAttributeType.imSearchBoth, imSearchEnabledDisabled.imSearchEnabledOnly, false);

                    for (int i = 1; i <= cX.Count; i++)
                    {
                        result.Add(cX.ItemByIndex(i));
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Fatal Error. Description: " + ex.Message, "GetValues", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        private void btnOK_Click_1(object sender, EventArgs e)
        {
            _selectCustom = GetSelectedValue();

            if (_selectCustom == null)
            {
                MessageBox.Show("Please select one Custom value", "No Custom selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                _selectCustomXTR = new XTR_Custom();
                _selectCustomXTR._WorkSiteOject = (_selectCustom);
                _isCancel = false;
                this.Close();
            }
        }

        private IManCustomAttribute GetSelectedValue()
        {
            if (dataGridViewCustom.SelectedRows.Count > 0)
            {
                try
                {
                    return _searchResult.Where(x => x.Name.Equals(dataGridViewCustom.SelectedRows[0].Cells[0].Value)).First();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error GetSelectedValue", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return null;
        }

        private void btnFindNow_Click_1(object sender, EventArgs e)
        {
            LoadValues(txtSearch.Text);
        }

        private void rdStart_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _isCancel = true;
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DoubleClickRecord(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                _selectCustom = GetSelectedValue();

                if (_selectCustom == null)
                {
                    MessageBox.Show("Please select one Custom value", "No Custom selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    _selectCustomXTR = new XTR_Custom();
                    _selectCustomXTR._WorkSiteOject = (_selectCustom);
                    _isCancel = false;
                    this.Close();
                }
            }

        }

        public void LoadAdminSession(string username, string password)
        {
            IMANADMIN.NRTDMS dms = new IMANADMIN.NRTDMS();
            IMANADMIN.NRTSession _a_session = dms.Sessions.Add(_sess.ServerName);

            _a_session.Login(username, password);

            if (!_a_session.Connected)
            {
                throw new Exception("Unable to connect to WorkSite using IManAdmin. Review your credencials and connection.");
            }

            _dataBasenameAdmin = (IMANADMIN.NRTDatabase)_a_session.Databases.Item(_dataBasename.Name);
        }

        string alias;
        string desc;
        private void btnAdd_Click(object sender, EventArgs e)
        {

            NewCustomXDlg newCustomX = new NewCustomXDlg(txtSearch.Text, desc);
            newCustomX.ShowDialog();

            if (newCustomX.IsSave)
            {
                alias = newCustomX.GetAlias();
                desc = newCustomX.GetDescription();

                txtSearch.Text = alias.Trim();

                try
                {
                    IMANADMIN.AttributeID tabela = XTRWorkSite.Helper.EnumBinding.GetAttributeID((XTRWorkSite.WSObjects.Enums.XTR_EnumAttributeID)(int)_CustomX);

                    if (string.IsNullOrEmpty(_parentValue))
                    {
                        _dataBasenameAdmin.InsertParentCustomField(tabela, alias.Trim(), desc, true);
                    }
                    else
                    {
                        _dataBasenameAdmin.InsertChildCustomField(tabela, _parentValue, alias.ToUpper().Trim(), desc, true);
                    }

                    desc = null;

                    btnFindNow_Click_1(null, null);
                }
                catch (Exception )
                {
                    MessageBox.Show("Error creating the item. Item already exist.", "Item error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    btnAdd_Click(sender, e);
                }
            }
        }
    }
}
