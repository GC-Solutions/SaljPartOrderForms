using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaljPartOrderForms
{
    public partial class inputdate : Form
    {
        private string theDate; // field

        public string TheDate   // property
        {
            get { return theDate; }   // get method
            set { theDate = value; }  // set method
        }

        public inputdate()
        {
            InitializeComponent();
            dtLevdate.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TheDate = dtLevdate.Text;
            if (!string.IsNullOrEmpty(TheDate)){
                DialogResult = DialogResult.OK;
                this.Close();
            }
        }


        private void txtDate_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button1_Click(null,null);
                e.Handled = true; //Handle the Keypress event (suppress the Beep)
            }
        }
    }
}
