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
using System.Threading;
using Hyland.Unity;


namespace OnBase_Export_Management
{
    public partial class Form1 : Form
    {
        bool isDisconnected = false;
        Dictionary<string, string> dicDTs = new Dictionary<string, string>();
        string error = string.Empty;
        string basePath = System.Configuration.ConfigurationManager.AppSettings["BasePath"].ToString();
        
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
                dicDTs.Clear();
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = checkBox1.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = true;
                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                lblUser.Text = "Welcome " + obc.RealName();
                button1.Text = "Connected";
                button1.Enabled = false;
                string[] dtgDT = File.ReadAllText("DTG-DT.txt").Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
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
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //for (int i = 1; i <= 100; i++)
            //{
            //    // Wait 50 milliseconds.  
            //    Thread.Sleep(50);
            //    // Report progress.  

            //    backgroundWorker1.ReportProgress(i);
            //}
        }
        private void backgroundWorker1_ProgressChanged(object sender,
        ProgressChangedEventArgs e)
        {
            //// Change the value of the ProgressBar  
            //pBar.Value = e.ProgressPercentage;
            //// Set the text.  
            //this.Text = "Progress: " + e.ProgressPercentage.ToString() + "%";
        }
        private void WriteToAppLogs(string line)
        {
            appLog.Items.Add(line);
            appLog.SelectedIndex = appLog.Items.Count - 1;
            appLog.SelectedIndex = -1;
            appLog.Refresh();

        }
        private void SetProgress(int done, int total)
        {
            label8.Text = "Progress " + done + "/" + total;
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            bool annotation = radioButton1.Checked;
            this.Cursor = Cursors.AppStarting;
            appLog.Items.Clear();
            WriteToAppLogs("Downlaod Started...");
            this.Cursor = Cursors.WaitCursor;
            string from = string.Empty;
            string to = string.Empty;
            long DHFrom = 0;
            long DHTo = 0;
            DateTime dtFrom = new DateTime();
            DateTime dtTo = new DateTime();
            string basePath = System.Configuration.ConfigurationManager.AppSettings["BasePath"].ToString();
            if ((txtDHFrom.Text == "") || (txtDHFrom.Text == string.Empty))
            {
                txtDHFrom.Text = "0";
            }
            if (txtDHTo.Text == "" || txtDHTo.Text == string.Empty)
                DHTo = long.MaxValue;
            DHFrom = Convert.ToInt64(txtDHFrom.Text);

            DHTo = Convert.ToInt64(txtDHTo.Text);
            if (DHTo == 0)
                DHTo = long.MaxValue;

            from = dtpFrom.Value.ToString("MM/dd/yyyy");
            to = dtpTo.Value.ToString("MM/dd/yyyy");
            dtFrom = Convert.ToDateTime(from);
            dtTo = Convert.ToDateTime(to);

            string docType = cmbDocType.SelectedItem.ToString().Trim();
            string dtg = cmbDocTypeGroup.SelectedItem.ToString().Trim();

            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = txtDHTo.Enabled = txtDHFrom.Enabled = false;
            if (checkBox1.Checked)
            {
                if (DHFrom > DHTo)
                {
                    MessageBox.Show("Invalid Document Handle Range ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                long total = DHTo - DHFrom + 1;
                int dc = 1;
                for (long i = DHFrom; i <= DHTo; i++)
                {
                    Document doc = obc.GetDocumentByIDs(i);
                    if (doc != null)
                    {
                        bool isDownloaded = false;
                        if (annotation)
                        {
                            isDownloaded = obc.SaveToDiscWithAnnotation(basePath, doc, true);
                        }
                        else
                        {
                            isDownloaded = obc.SaveToDiscWithoutAnnotation(basePath, doc);
                        }

                        if (isDownloaded)
                        {
                            WriteToAppLogs("Document Downloaded - Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);

                        }
                        else
                        {
                            WriteToAppLogs("Failed to Download document with Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);

                        }
                    }
                    else
                    {
                        WriteToAppLogs("Document not found with Document Handle " + i );
                    }
                    SetProgress(dc, (int)total);
                    dc++;

                }

                txtDHTo.Enabled = txtDHFrom.Enabled = true;
            }
            else
            {
                List<Document> docList = new List<Document>();
                if (dtpFrom.Value > dtpTo.Value)
                {
                    WriteToAppLogs("Invalid Date Range Selected...");
                    MessageBox.Show("Please select valid date range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = true;
                    return;
                }
                if(docType.ToUpper() == "ALL")
                {
                    if (dtg == null || dtg == string.Empty)
                    {
                        this.Cursor = Cursors.Default;
                        MessageBox.Show("Please Select Document Type Group", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    docList = obc.GetDocumentListByDocumentTypeGroup(dtg, dtFrom, dtTo, DHFrom, DHTo);
                }
                else
                {
                    docList = obc.GetDocumentList(docType, dtFrom, dtTo, DHFrom, DHTo);
                }
                //docList = docList.SingleOrDefault(()=>)
                //Comparison<Document> mycomp = (x, y) => x.ID.CompareTo(y.ID);
                docList.Sort((x, y) => x.ID.CompareTo(y.ID));
                //docList = docList.Sort(x => x.ID);
                WriteToAppLogs("Total Document are " + docList.Count);
                WriteToAppLogs("Total Estimated time " + Math.Floor((docList.Count * 3.0) / 60) + " Minutes and " + ((docList.Count * 3.0) % 60) + " Seconds");
                DateTime dtStart = System.DateTime.Now;
                WriteToAppLogs("Downlaod Start Time " + dtStart.ToString("MM-dd-yyyy HH:mm:ss"));
                int counter = 1;
                foreach (Document doc in docList)
                {

                    if (obc.SaveToDiscWithoutAnnotation(basePath, doc))
                    {
                        WriteToAppLogs("Document Downloaded - Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);

                    }
                    else
                    {
                        WriteToAppLogs("Failed to Download document with Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);

                    }
                    SetProgress(counter, docList.Count);
                    counter++;
                }
                DateTime dtEnd = System.DateTime.Now;
                WriteToAppLogs("Downlaod Start Time " + dtEnd.ToString("MM-dd-yyyy HH:mm:ss"));
                WriteToAppLogs("Actual Time in downloads = " + (dtEnd - dtStart).Seconds + " Seconds");
                WriteToAppLogs("Downlaod Finished...");                
                
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = txtDHFrom.Enabled=txtDHTo.Enabled = btnDisconnect.Enabled = true;

                MessageBox.Show("Documents Downloaded successfully for Document Type " + docType + " and date range from " + from + " to " + to);
            }
            this.Cursor = Cursors.Default;
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            DateTime releaseDate = new DateTime(2022, 10, 31);
            if (System.DateTime.Now < releaseDate)
            {
                MessageBox.Show("System Date is incorrect","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            string licenseKey = System.Configuration.ConfigurationManager.AppSettings["LicenseKey"].ToString();
            DateTime validTill = new DateTime();
            if (licenseKey == "5FF34-678E1-01012-078AC")
            {
                validTill = releaseDate.AddDays(11);
            }
            else if (licenseKey == "0FA12-6A8B0-10120-058AD")
            {
                validTill = releaseDate.AddDays(18);
            }
            else if (licenseKey == "2EF34-628A0-105B0-9A8AC")
            {
                validTill = releaseDate.AddDays(25);
            }
            else if (licenseKey == "1EE11-05EAC-199A0-98ABC")
            {
                validTill = releaseDate.AddDays(55);
            }
            else if (licenseKey == "1ED11-05EAC-197A0-98AEC")
            {
                validTill = releaseDate.AddYears(10);
            }
            else
            {
                MessageBox.Show("Please provide valid license Key","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            //MessageBox.Show("Valid Till "+validTill.ToString("MM-dd-yyyy"));

            if (System.DateTime.Now > validTill)
            {
                MessageBox.Show("License Expired or Invalid Key! Please contact to the administrator", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            
            
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
            btnDisconnect.Enabled = lblUser.Visible = false;
            button1.Text = "Connect";
            cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = false;
        }

        private void txtDHFrom_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            //if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            //{
            //    e.Handled = true;
            //}
        }

        private void txtDHTo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            //if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            //{
            //    e.Handled = true;
            //}
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = false;
                txtDHFrom.Enabled = txtDHTo.Enabled = true;
            }
            else
            {
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = true;
                txtDHTo.Text = txtDHFrom.Text = "0";
                //txtDHFrom.Enabled = txtDHTo.Enabled = false;
            }
        }
    }
}
