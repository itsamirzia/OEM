using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OnBase_Export_Management
{
    public partial class login : Form
    {
        public login()
        {
            InitializeComponent();
        }

        private void login_Load(object sender, EventArgs e)
        {

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            
            if (txtUsername.Text.ToUpper() != "ADMINISTRATOR")
            {
                MessageBox.Show("Invalid Username");
                return;
            }
            if (txtPassword.Text != "Datum22@123!")
            {
                MessageBox.Show("Invalid Password");
                return;
            }
            this.Hide();
            Form1 form = new Form1();
            form.ShowDialog();
            this.Close();
        }
    }
}
