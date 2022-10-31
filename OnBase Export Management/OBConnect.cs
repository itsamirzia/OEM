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
using System.Configuration;


namespace OnBase_Export_Management
{
    public partial class OBConnect : Form
    {
        bool isConnect = false;
        bool NTAuth = false;
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
                if (NTAuth)
                {
                    isConnect = obc.Connect(txtAppURL.Text.ToString(), txtDataSource.Text.ToString(), txtUsername.Text.ToString(), txtPassword.Text.ToString(),true);
                }
                else
                {
                    isConnect = obc.Connect(txtAppURL.Text.ToString(), txtDataSource.Text.ToString(), txtUsername.Text.ToString(), txtPassword.Text.ToString());
                }

                if (isConnect)
                {
                    MessageBox.Show("User connection successful","",MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void OBConnect_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            string OBConn = System.Configuration.ConfigurationManager.AppSettings["OBConnString"].ToString();
            string[] loginArray = OBConn.Split(';');
            foreach (string str in loginArray)
            {
                string[] keyVal = str.Split('=');
                string key = keyVal[0].Trim().ToString();
                string val = keyVal[1].Trim().ToString();
                if (key.ToUpper() == "APPURL")
                    txtAppURL.Text = val;
                if (key.ToUpper() == "DATASOURCE")
                    txtDataSource.Text = val;
                if (key.ToUpper() == "USERNAME")
                    txtUsername.Text = val;
                if (key.ToUpper() == "PASSWORD")
                    txtPassword.Text = val;
            }
            
        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem.ToString().ToUpper() == "NT AUTHENTICATION")
            {
                txtUsername.Enabled = txtPassword.Enabled = false;
                NTAuth = true;
            }
            else
            {
                txtUsername.Enabled = txtPassword.Enabled = true;
                NTAuth = false;
            }
        }
    }
}
