using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace XTRWorkSite.CustomForms
{
    public partial class LoginDlg : Form
    {
        public bool _isCancel = true;
        public WorkSiteAccess _wkAccess = null;

        public LoginDlg()
        {
            InitializeComponent();
        }

        public LoginDlg(string host, string bd, string username, string password)
        {
            InitializeComponent();
            txtServer.Text = host;
            txtDatabase.Text = bd;
            txtUsername.Text = username;
            txtPassword.Text = password;   
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtServer.Text) || string.IsNullOrEmpty(txtDatabase.Text))
                {
                    MessageBox.Show("Server e DataBase are required", "Required Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (chkTrusted.Checked || (!string.IsNullOrEmpty(txtUsername.Text) && (!string.IsNullOrEmpty(txtPassword.Text))))
                {
                    try
                    {
                        _wkAccess = new WorkSiteAccess(txtServer.Text, txtUsername.Text, txtPassword.Text, txtDatabase.Text, chkTrusted.Checked);

                        if (_wkAccess.CheckIManageConnection())
                        {
                            _isCancel = false;
                            Close();
                        }
                        else
                        {
                            MessageBox.Show("Data invalid", "Data invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Auth Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _isCancel = true;
            this.Close();
        }

        private void chkTrusted_CheckedChanged(object sender, EventArgs e)
        {
            if (chkTrusted.Checked)
            {
                txtPassword.Enabled = false;
                txtUsername.Enabled = false;
            }
            else
            {
                txtPassword.Enabled = true;
                txtUsername.Enabled = true;
            }
        }
    }
}
