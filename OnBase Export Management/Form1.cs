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
    public partial class Form1 : Form
    {
        string error = string.Empty;
        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OBConnect ob = new OBConnect();
            ob.ShowDialog();
            if (!ob.IsConnected())
            {
                this.Close();
            }
            else
            {
                
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = true;
                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                lblUser.Text = "Welcome " + obc.RealName();
                button1.Text = "Connected";
                List<string> dtgL = obc.GetDocumentTypeGroupList(ref error);
                cmbDocTypeGroup.DataSource = dtgL;
                
                
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedValue = cmbDocTypeGroup.SelectedItem.ToString();
            long id = Convert.ToInt64( selectedValue.Split(new[] { "---" }, StringSplitOptions.RemoveEmptyEntries)[0].Trim());
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();

            List<string> dtL = obc.GetDocumentTypeList(id, ref error);
            cmbDocType.DataSource = dtL;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            if (obc.ExportDocument(@"C:\Users\mohdziya\Desktop\Ziya\Export Test", cmbDocType.SelectedItem.ToString().Split(new[] { "---" }, StringSplitOptions.RemoveEmptyEntries)[1].Trim(), dtpFrom.Value, dtpTo.Value, false))
            {
                MessageBox.Show("Exported Successfully");
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            obc.Disconnect();
        }
    }
}
