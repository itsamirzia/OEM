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
using System.IO;

namespace OnBase_Export_Management
{
    public partial class Form1 : Form
    {
        bool isDisconnected = false;
        Dictionary<string, string> dicDTs = new Dictionary<string, string>();
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
                
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = true;
                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                lblUser.Text = "Welcome " + obc.RealName();
                button1.Text = "Connected";
                button1.Enabled = false;
                string[] dtgDT = File.ReadAllText("DTG-DT.txt").Split(new[] { "**********" }, StringSplitOptions.RemoveEmptyEntries);
                List<string> DTGs = new List<string>();
                List<string> DTs = new List<string>();
                DTs.Add("ALL");
                foreach (string dts in dtgDT)
                {
                    
                    string[] temp1 = dts.Split('|');
                    dicDTs.Add(temp1[0].Trim(), temp1[1].Trim());
                    DTGs.Add(temp1[0].Trim());
                    string[] temp2 = temp1[1].Trim().Split(',');
                    foreach(string dt in temp2)
                        DTs.Add(dt);

                }
                cmbDocTypeGroup.DataSource = DTGs;
                cmbDocType.DataSource = DTs;
                this.Refresh();

            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbDocType.DataSource = null;
            string selectedValue = cmbDocTypeGroup.SelectedItem.ToString();
            //long id = Convert.ToInt64( selectedValue.Split(new[] { "---" }, StringSplitOptions.RemoveEmptyEntries)[0].Trim());
            //OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            //List<string> dtL = obc.GetDocumentTypeList(id, ref error);
            List<string> DTs = new List<string>();
            string[] dts = dicDTs[selectedValue].Split(',');
            foreach (string dt in dts)
                DTs.Add(dt);
            cmbDocType.DataSource = DTs;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = false;
            if (dtpFrom.Value > dtpTo.Value)
            {
                MessageBox.Show("Please select valid date range","Error",MessageBoxButtons.OK, MessageBoxIcon.Stop);
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = true;
                return;
            }
            int docCounter = 0;
            string from = dtpFrom.Value.ToString("MM/dd/yyyy");
            string to = dtpTo.Value.ToString("MM/dd/yyyy");
            DateTime dtFrom = Convert.ToDateTime(from);
            DateTime dtTo = Convert.ToDateTime(to);
            int totalDays = (dtTo - dtFrom).Days +1;
            string basePath = ConfigurationSettings.AppSettings["BasePath"].ToString();
            string docType = cmbDocType.SelectedItem.ToString().Trim();
            int counter = 0;
            while (dtFrom <= dtTo)
            {
                DateTime dt1 = Convert.ToDateTime(dtFrom.ToString("MM/dd/yyyy 00:00:00"));
                DateTime dt2 = Convert.ToDateTime(dtFrom.ToString("MM/dd/yyyy 11:59:59"));
                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                int docCount = obc.ExportDocument(basePath, docType, dt1, dt2, false);
                if (docCount > 0)
                {
                    File.AppendAllText(basePath + "\\" + from + "_" + to + ".txt", "Total doc Exported for doctype = "+docType+" and count = "+docCount);
                    docCounter += docCount;
                }
                dtFrom.AddDays(1);
                counter++;
            }
            cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = true;
            MessageBox.Show(docCounter +" Document downloaded successfully for Document Type "+docType+ " and date range from "+from+" to "+to);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            if(!isDisconnected)
                obc.Disconnect();
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            obc.Disconnect();
            button1.Enabled = isDisconnected = true;
            btnDisconnect.Enabled = false;
        }
    }
}
