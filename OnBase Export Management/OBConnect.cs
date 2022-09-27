using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OBConnector;


namespace OnBase_Export_Management
{
    public partial class OBConnect : Form
    {
        bool isConnect = false;
        public OBConnect()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                if (obc.Connect(txtAppURL.Text.ToString(), txtDataSource.Text.ToString(), txtUsername.Text.ToString(), txtPassword.Text.ToString()))
                {
                    MessageBox.Show("User connection successful ");
                    this.isConnect = true;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                rtbError.Text = ex.Message;
                rtbError.Visible = true;
                isConnect = false;
            }
        }
        public bool IsConnected()
        {
            return isConnect;
        }

        
    }
}
