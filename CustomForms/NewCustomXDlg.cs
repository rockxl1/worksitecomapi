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
    public partial class NewCustomXDlg : Form
    {
        public bool IsSave = false;

        public NewCustomXDlg(string title, string desc)
        {
            InitializeComponent();
            txtAlias.ShortcutsEnabled = false;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrWhiteSpace(txtAlias.Text) || string.IsNullOrWhiteSpace(txtDescription.Text))
            {
                IsSave = false;
                MessageBox.Show("All fields are required", "Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                IsSave = true;
                this.Close();
            } 
        }

        public string GetAlias()
        {
            return txtAlias.Text;
        }

        public string GetDescription()
        {
            return txtDescription.Text;
        }

        private void txtAlias_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(txtAlias.TextLength<=32)
            {
                if(IsCharValid(e.KeyChar) || e.KeyChar == '\b')
                {
                    e.KeyChar = char.ToUpper(e.KeyChar);
                }
                else
                {
                    e.Handled = true;
                }
            }
        }

        public bool IsCharValid(char letter)
        {
            string numbers = "32|38|40-41|44-57|65-91|93|95|97-122|123|125|128|130-144|148-154|160-165|170|186|181-183|192-196|198-214|217-220|224-228|229|231-239|242-246|249-252";
            List<int> asciiValidos = DecompemSyntaxEmNumeros(numbers);

            if(letter == ' ')
            {
                return false;
            }

           int ascii = (int)letter;

            if (!asciiValidos.Contains(ascii))
            {
                return false;
            }
            else
            {
                return true;
            }

        }

        public static List<int> DecompemSyntaxEmNumeros(string regrasLoggic)
        {
            // "65-90|97-122|123|125";
            List<int> lista = new List<int>();
            string[] grupos = regrasLoggic.Split('|');

            foreach (var item in grupos)
            {
                if (item.Contains("-"))
                {
                    string[] inicioFim = item.Split('-');
                    for (int i = Convert.ToInt32(inicioFim[0]); i <= Convert.ToInt32(inicioFim[1]); i++)
                    {
                        lista.Add(i);
                    }
                }
                else
                {
                    lista.Add(Convert.ToInt32(item));
                }
            }

            return lista;
        }

        private void txtDescription_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtAlias.TextLength > 256)
            {
                e.Handled = true;
            }
          
        }

        private void txtAlias_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
