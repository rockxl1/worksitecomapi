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
    public partial class ClassDlg : Form
    {
        public bool _isCancel = true;

        public IManDocumentClass _selectClass;
        public XTR_DocumentClass _selectClassXTR;

        private IManDMS _m_dms { get; set; }
        private IManSession _sess { get; set; }
        private IManDatabase _dataBasename { get; set; }

        private List<XTR_DocumentClass> classes;


        public ClassDlg(WorkSiteAccess worksite, string database = null)
        {
            InitializeComponent();
            _m_dms = worksite._m_dms;
            _sess = worksite._m_session;
            _dataBasename = worksite._m_dataBase.First();

            if (!string.IsNullOrEmpty(database))
            {
                _dataBasename = worksite._m_dataBase.Where(x => x.Name.ToUpper().Equals(database.ToUpper())).First();
            }

            InitializeGrid();
        }

        public ClassDlg(IManSession session, IManDMS dms, IManDatabase dataBase)
        {
            _sess = session;
            _dataBasename = dataBase;
            _m_dms = dms;
            InitializeComponent();
            InitializeGrid();
        }

        public ClassDlg(string host, string username, string password, string database, bool trusted)
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

            _dataBasename = (IManage.IManDatabase)_sess.Databases.ItemByName(database);
            _sess.Timeout = 30;


            InitializeComponent();
            InitializeGrid();
        }

        private void InitializeGrid()
        {
            _sess.MaxRowsForSearch = 9999;
            _sess.MaxRowsNonSearch = 9999;

            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.CellDoubleClick += Mouse_doubleClick;
            dataGridView1.Columns.Add("Alias", "Alias");
            dataGridView1.Columns.Add("Descricao", "Descrição");
            LoadValues(string.Empty);

            txtClass.KeyDown += new KeyEventHandler(descp_KeyDown);
        }

        private void descp_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                LoadValues(txtClass.Text);
            }
        }

        public void LoadValues(string query)
        {
            try
            {
                dataGridView1.Rows.Clear();

                if (classes == null || classes.Count == 0)
                {
                    classes = new List<XTR_DocumentClass>();

                    IManDocumentClasses docClass = _dataBasename.SearchDocumentClasses(string.Empty, imSearchAttributeType.imSearchBoth, true);

                    for (int i = 1; i <= docClass.Count; i++)
                    {
                        IManDocumentClass item = docClass.ItemByIndex(i);
                        XTR_DocumentClass docCl = new XTR_DocumentClass();
                        docCl._WorkSiteOject = (item);
                        classes.Add(docCl);
                    }
                }

                var result = classes;

                if (!string.IsNullOrEmpty(query))
                {
                    result = result.Where(x => x.Description.ToUpper().Contains(query.ToUpper()) ||
                        x.Name.Contains(query.ToUpper())).ToList();
                }


                foreach (XTR_DocumentClass item in result)
                {
                    dataGridView1.Rows.Add(item.Name, item.Description);
                }
                lblItemsCount.Text = result.Count.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void Mouse_doubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex != -1)
            { 
            if (dataGridView1.SelectedRows.Count > 0)
            {
                try
                {
                    _selectClassXTR = classes.Where(x => x.Name.Equals(dataGridView1.SelectedRows[0].Cells[0].Value)).First();
                    _selectClass = _selectClassXTR._WorkSiteOject;
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


        private void btnOk_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                try
                {
                    _selectClassXTR = classes.Where(x => x.Name.Equals(dataGridView1.SelectedRows[0].Cells[0].Value)).First();
                    _selectClass = _selectClassXTR._WorkSiteOject;
                    _isCancel = false;
                    Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error. Description: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            LoadValues(txtClass.Text);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _isCancel = true;
            this.Close();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void lblItemsCount_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void txtClass_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
