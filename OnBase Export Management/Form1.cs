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
using System.Text.RegularExpressions;

namespace OnBase_Export_Management
{
    public partial class Form1 : Form
    {
        bool isDisconnected = true;
        Dictionary<string, string> dicDTs = new Dictionary<string, string>();
        string error = string.Empty;
        string basePath = System.Configuration.ConfigurationManager.AppSettings["BasePath"].ToString();
        bool metadataXML = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["MetadataXML"].ToString());
        string uniqueID = string.Empty;
        int docsInBatch = 0;
        
        Dictionary<string, string> ctxDocsDic = new Dictionary<string, string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OBConnect ob = new OBConnect();
                ob.ShowDialog();
                if (!ob.IsConnected())
                {
                    //this.Close();
                }
                else
                {
                    isDisconnected = false;

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
                        foreach (string dt in temp2)
                            DTs.Add(dt);

                    }
                    cmbDocTypeGroup.DataSource = DTGs;
                    cmbDocType.DataSource = DTs;
                    this.Refresh();

                }
            }
            catch
            {
                this.Dispose();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbDocType.DataSource = null;
            string selectedValue = cmbDocTypeGroup.SelectedItem.ToString();
            List<string> DTs = new List<string>();
            DTs.Add("ALL");
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
            if(Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["EnableLogsInDB"].ToString()))
                db.ExecuteNonQuery("insert into [dbo].Logs values ('"+uniqueID+"','"+line+"',GETDATE());");
            appLog.Items.Add(line);
            appLog.SelectedIndex = appLog.Items.Count - 1;
            appLog.SelectedIndex = -1;
            appLog.Refresh();

        }
        private void SetProgress(int done, int total)
        {
            label8.Text = "Progress " + done + "/" + total;
        }
        private void InertExceptionIfFound(long docID, string uID)
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            if (obc.CurrentException() != string.Empty)
            {
                db.ExecuteNonQuery("insert into [dbo].[Exception] values (" + docID + ",'" + uniqueID + "','" + obc.CurrentException() + "','',GETDATE());");
            }

        }
        private Dictionary<string, string> ConvertDTtoDict(DataTable dt)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (DataRow dr in dt.Rows)
            {
                dict.Add(dr[0].ToString(), dr[1].ToString());
            }
            return dict;
        }
        private List<string> ConvertDTtoList(DataTable dt)
        {
            List<string> dict = new List<string>();
            foreach (DataRow dr in dt.Rows)
            {
                dict.Add(dr[0].ToString());
            }
            return dict;
        }
        private bool IfFoundInDic(Dictionary<string,string>dict, string docType)
        {
            bool isFound = false;
            try
            {
                foreach (KeyValuePair<string, string> kvp in dict)
                {
                    if(kvp.Key == docType)
                    {
                        isFound = true;
                    }
                }

            }
            catch
            {

            }
            return isFound;
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (LicenseCheck() != 0)
                {
                    WriteToAppLogs("License Expired or Invalid Key");
                    MessageBox.Show("License Expired or Invalid Key", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                DataTable dtOBCtxDT = new DataTable();// db.ExecuteSQLQuery("SELECT trim(CTXDt),trim(ReserveKey)  FROM [dbo].[OBCTXDt]");
                ctxDocsDic = ConvertDTtoDict(dtOBCtxDT);
                docsInBatch = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DocsInBatch"].ToString());
                bool annotation = radioButton1.Checked;
                this.Cursor = Cursors.AppStarting;
                appLog.Items.Clear();
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

                string docType = string.Empty;
                try
                {
                    docType = cmbDocType.SelectedItem.ToString().Trim();
                }
                catch
                {
                    MessageBox.Show("Please select the document type to proceed further");
                    return;
                }
                string dtg = string.Empty;
                try
                {
                    dtg = cmbDocTypeGroup.SelectedItem.ToString().Trim();
                }
                catch 
                {
                    MessageBox.Show("Please select the valid document type group");
                    return;
                }

                this.Cursor = Cursors.WaitCursor;

                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = txtDHTo.Enabled = txtDHFrom.Enabled = false;

                uniqueID = GetUniqueID();
                WriteToAppLogs("Preparing Data. Please wait...");
                DataTable dtOBDT = db.ExecuteSQLQuery("SELECT trim([OBDocType]),trim([ALFDocType]) FROM [dbo].[OBDocTypeVsALFDocType]");
                obc.SetOBDTvsALFDT(ConvertDTtoDict(dtOBDT));
                DataTable dtOBKey = db.ExecuteSQLQuery("select trim(DocumentType)+'_'+trim(OBKey), alfkey from [dbo].[OBKeyVsALFKey]");
                obc.SetOBKeyvsALFKey(ConvertDTtoDict(dtOBKey));
                DataTable dtOBPath = db.ExecuteSQLQuery("SELECT trim([OBDTG]),[DownloadPath] FROM [dbo].[OBDTGVsPath]");
                obc.SetOBDTGvsPath(ConvertDTtoDict(dtOBPath));
                

                if (checkBox1.Checked)
                {
                    if (DHFrom > DHTo)
                    {
                        MessageBox.Show("Invalid Document Handle Range ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    long total = DHTo - DHFrom + 1;
                    if (chkWSFeed.Checked)
                    {
                        if (db.ExecuteNonQuery("insert into [dbo].[SearchLog] values (" + uniqueID + ",'','',''," + DHFrom + "," + DHTo + ",'No',GETDATE()," + total + ",0,'','WS-Pending','" + obc.RealName() + "');"))
                        {
                            MessageBox.Show("Feed is added for window service");
                            WriteToAppLogs("Feed is added for window service");
                        }
                        else
                        {
                            MessageBox.Show("Something is wrong with the database. Not able to set the feed for window service.");
                            WriteToAppLogs("Something is wrong with the database. Not able to set the feed for window service.");
                        }
                        return;
                    }
                    db.ExecuteNonQuery("insert into [dbo].[SearchLog] values (" + uniqueID + ",'','',''," + DHFrom + "," + DHTo + ",'No',GETDATE()," + total + ",0,'','Pending','" + obc.RealName() + "');");
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
                                isDownloaded = obc.SaveToDiscWithoutAnnotation(basePath, doc, uniqueID);
                            }

                            if (isDownloaded)
                            {
                                WriteToAppLogs("Document Downloaded - Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);
                                db.ExecuteNonQuery("insert into [dbo].[DownloadedItems] values (" + doc.ID + ",'" + basePath + "\\" + doc.DocumentType.Name + "',GETDATE(),'Success','" + uniqueID + "');");
                                db.ExecuteNonQuery("update table [dbo].[SearchLogs] set LastExecutedDH=" + doc.ID + " where SearchID='" + uniqueID + "');");
                                if (obc.CurrentException() != string.Empty)
                                    InertExceptionIfFound(doc.ID, uniqueID);


                            }
                            else
                            {
                                WriteToAppLogs("Failed to Download document with Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);
                                db.ExecuteNonQuery("insert into [dbo].[DownloadedItems] values (" + doc.ID + ",'" + basePath + "\\" + doc.DocumentType.Name + "',GETDATE(),'Failed','" + uniqueID + "');");
                                if (obc.CurrentException() != string.Empty)
                                    InertExceptionIfFound(doc.ID, uniqueID);
                                //db.ExecuteNonQuery("insert into [dbo].[Exception] values (" + doc.ID + ",'" + uniqueID + "','" + obc.CurrentException() + "','',GETDATE());");


                            }
                        }
                        else
                        {
                            WriteToAppLogs("Document not found with Document Handle " + i);
                        }
                        SetProgress(dc, (int)total);
                        dc++;

                    }

                    txtDHTo.Enabled = txtDHFrom.Enabled = true;
                }
                else
                {
                    if (dtpFrom.Value > dtpTo.Value)
                    {
                        WriteToAppLogs("Invalid Date Range Selected...");
                        MessageBox.Show("Please select valid date range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = true;
                        return;
                    }
                    List<string> documentTypeList = new List<string>();

                    List<Document> docList = new List<Document>();

                    string isDTG = "No";
                    int dtgDocCount = 0;
                    if (docType.ToUpper() == "ALL")
                    {
                        if (dtg == null || dtg == string.Empty)
                        {
                            this.Cursor = Cursors.Default;
                            MessageBox.Show("Please Select Document Type Group", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        string docTypes = dicDTs[dtg].ToString();
                        if (chkWSFeed.Checked)
                        {
                            if (db.ExecuteNonQuery("insert into [dbo].[SearchLog] values (" + uniqueID + ",'" + docTypes + "','" + from + "','" + to + "'," + DHFrom + "," + DHTo + ",'No',GETDATE(),0,'0','','WS-Pending','" + obc.RealName() + "');"))
                            {
                                MessageBox.Show("Feed is added for window service for doctype "+docTypes);
                                WriteToAppLogs("Feed is added for window service for doctype " + docTypes);
                            }
                            else
                            {
                                MessageBox.Show("Not able to set the feed for window service.");
                                WriteToAppLogs("Not able to set the feed for window service.");
                            }
                            return;
                        }
                        db.ExecuteNonQuery("insert into [dbo].[SearchLog] values (" + uniqueID + ",'" + docTypes + "','" + from + "','" + to + "'," + DHFrom + "," + DHTo + ",'No',GETDATE(),0,'0','','WS-Pending','" + obc.RealName() + "');");

                        string[] dtgDocTypes = docTypes.Trim().Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string dtgDTType in dtgDocTypes)
                        {
                            documentTypeList.Add(dtgDTType.Trim());
                        }
                        isDTG = "Yes";
                    }
                    else
                    {
                        if (chkWSFeed.Checked)
                        {
                            if (db.ExecuteNonQuery("insert into [dbo].[SearchLog] values (" + uniqueID + ",'" + docType + "','" + from + "','" + to + "'," + DHFrom + "," + DHTo + ",'No',GETDATE(),0,'0','','WS-Pending','" + obc.RealName() + "');"))
                            {
                                MessageBox.Show("Feed is added for window service for doctype " + docType);
                                WriteToAppLogs("Feed is added for window service for doctype " + docType);
                            }
                            else
                            {
                                MessageBox.Show("Not able to set the feed for window service.");
                                WriteToAppLogs("Not able to set the feed for window service.");
                            }
                            return;
                        }
                        documentTypeList.Add(docType);

                    }
                    WriteToAppLogs("Download Process Started...");

                    foreach (string documentType in documentTypeList)
                    {
                        WriteToAppLogs("Retrieving Document for " + documentType + ". It may take longer time. Please wait...");
                        docList.Clear();
                        docList = obc.GetDocumentList(documentType, dtFrom, dtTo, DHFrom, DHTo);
                        InertExceptionIfFound(0, uniqueID);
                        if (docList == null)
                        {
                            if (isDTG.ToUpper() == "NO")
                            {
                                MessageBox.Show("Something is wrong in data. Check with Exception logs");
                                this.Cursor = Cursors.Default;
                                return;
                            }
                            else
                            {
                                WriteToAppLogs("Something is wrong in data. Check with Exception logs");
                                continue;
                            }
                        }
                        if (docList.Count == 0)
                        {
                            if (isDTG.ToUpper() == "NO")
                            {
                                MessageBox.Show("No document found in OnBase");
                                this.Cursor = Cursors.Default;
                                return;
                            }
                            else
                            {
                                WriteToAppLogs("No Document found for the document type " + documentType + " or document type not found in OnBase.");
                                continue;
                            }
                        }

                        if (isDTG.ToUpper() == "NO")
                            db.ExecuteNonQuery("insert into [dbo].[SearchLog] values (" + uniqueID + ",'" + documentType + "','" + from + "','" + to + "'," + DHFrom + "," + DHTo + ",'No',GETDATE()," + docList.Count + ",'0','','Pending','" + obc.RealName() + "');");
                        else
                            dtgDocCount += docList.Count;

                        docList.Sort((x, y) => x.ID.CompareTo(y.ID));
                        WriteToAppLogs("Total Document for document type " + documentType + " are " + docList.Count);
                        WriteToAppLogs("Total Estimated time for " + documentType + " " + Math.Floor((docList.Count * 3.0) / 60) + " Minutes and " + ((docList.Count * 3.0) % 60) + " Seconds");
                        DateTime dtStart = System.DateTime.Now;
                        WriteToAppLogs("Downlaod Start Time " + dtStart.ToString("MM-dd-yyyy HH:mm:ss"));

                        if (DownloadDocument(docList, basePath, metadataXML))
                        {

                            DateTime dtEnd = System.DateTime.Now;
                            WriteToAppLogs("Downlaod End Time for " + documentType + " is " + dtEnd.ToString("MM-dd-yyyy HH:mm:ss"));
                            WriteToAppLogs("Actual Time in downloads " + documentType + " are " + (dtEnd - dtStart).Seconds + " Seconds");
                            WriteToAppLogs("Downlaod Finished for " + documentType);

                            if (isDTG.ToUpper() == "NO")
                            {
                                db.ExecuteNonQuery("update [dbo].[SearchLog] set end_timestamp=GETDATE(), status='Complete' where SearchID='" + uniqueID + "';");
                                MessageBox.Show("Documents Downloaded successfully for Document Type " + docType + " and date range from " + from + " to " + to);
                            }
                        }
                    }
                    if (isDTG.ToUpper() == "YES")
                    {
                        db.ExecuteNonQuery("update [dbo].[SearchLog] set DocCount= " + dtgDocCount + ",end_timestamp=GETDATE(), status='Complete' where SearchID = '" + uniqueID + "'");
                        MessageBox.Show("Documents Downloaded successfully for Document Type group " + dtg + " and date range from " + from + " to " + to);
                    }

                }
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                this.Cursor = Cursors.Default;
            }
            finally
            {
                cmbDocTypeGroup.Enabled = cmbDocType.Enabled = dtpFrom.Enabled = dtpTo.Enabled = btnExport.Enabled = lblUser.Visible = btnDisconnect.Enabled = txtDHTo.Enabled = txtDHFrom.Enabled = true;
                this.Cursor = Cursors.Default;
            }
            
        }
        private bool DownloadDocument(List<Document> docList, string basePath, bool metadataXML)
        {
            try
            {
                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                int counter = 1;
                foreach (Document doc in docList)
                {
                    if (counter > docsInBatch)
                    {
                        string tempID = uniqueID;
                        uniqueID = GetUniqueID();
                        string queryInsertNewRecord = @"insert into [OB_Extraction].[dbo].[SearchLog] SELECT  '"+uniqueID+"',[DT_DTG],[DateRangeFrom],[DateRangeTo],[StartDocHandle],[EndDocHandle],[IsDTG],[run_timestamp],[DocCount],[LastExecutedDH],[end_timestamp],'InProgress',[RunByUserName] FROM [OB_Extraction].[dbo].[SearchLog] where SearchID='"+tempID+"'";
                        string queryUpdate = "update[dbo].[SearchLog] set Status = 'Complete',end_timestamp=GETDATE(), DocCount='"+ docsInBatch + "' where SearchID = '" + tempID + "'";
                        db.ExecuteNonQuery(queryInsertNewRecord);
                        db.ExecuteNonQuery(queryUpdate);
                        counter = 1;
                    }
                    bool downloadStatus = false;
                    if (IfFoundInDic(ctxDocsDic, doc.DocumentType.Name))
                    {
                        if (obc.SaveTXTToDiscWithoutAnnotation(basePath, doc, uniqueID, metadataXML))
                        {
                            downloadStatus = true;
                        }
                    }
                    else
                    {
                        if (obc.SaveToDiscWithoutAnnotation(basePath, doc, uniqueID, metadataXML))
                        {
                            downloadStatus = true;
                        }
                    }
                    if (downloadStatus)
                    {
                        WriteToAppLogs("Document Downloaded - Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);
                        if (obc.CurrentException() == string.Empty)
                        {
                            db.ExecuteNonQuery("insert into [dbo].[DownloadedItems] values (" + doc.ID + ",'" + basePath + "\\" + doc.DocumentType.Name + "',GETDATE(),'Success'," + uniqueID + ");");
                            db.ExecuteNonQuery("update [dbo].[SearchLog] set LastExecutedDH=" + doc.ID + ", DocCount='" + counter + "' where SearchID='" + uniqueID + "';");
                        }
                        else
                        {
                            db.ExecuteNonQuery("insert into [dbo].[Exception] values (" + doc.ID + ",'" + uniqueID + "','" + obc.CurrentException() + "','',GETDATE());");

                        }
                    }
                    else
                    {
                        WriteToAppLogs("Failed to Download document with Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);

                        db.ExecuteNonQuery("insert into [dbo].[DownloadedItems] values (" + doc.ID + ",'" + basePath + "\\" + uniqueID + "',GETDATE(),'Failed');");


                    }
                    counter++;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private string GetUniqueID()
        {
            var temp = Guid.NewGuid().ToString().Replace("-", string.Empty);
            var id = Regex.Replace(temp, "[a-zA-Z]", string.Empty).Substring(0, 12);
            if(db.HasDataRows("select * from [dbo].[SearchLog] where searchid='" + id + "'"))
            {
                GetUniqueID();
            }
            return id;
        }
        private int LicenseCheck()
        {
            DateTime releaseDate = new DateTime(2022, 12, 19);
            if (System.DateTime.Now < releaseDate)
            {
                MessageBox.Show("System Date is incorrect", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            string licenseKey = System.Configuration.ConfigurationManager.AppSettings["LicenseKey"].ToString();
            DateTime validTill = new DateTime();
            if (licenseKey == "5FF34-678E1-01012-078AC")
            {
                validTill = releaseDate.AddDays(5);
            }
            else if (licenseKey == "FFF12-6A8B0-10120-058AD")
            {
                validTill = releaseDate.AddDays(30);
            }
            else if (licenseKey == "FBF34-626A0-105B0-9A8AC")
            {
                validTill = releaseDate.AddDays(45);
            }
            else if (licenseKey == "ABCE11-05AAC-199A0-98ABC")
            {
                validTill = releaseDate.AddDays(60);
            }
            else if (licenseKey == "ABCDE-05EAC-197A0-98AEC")
            {
                validTill = releaseDate.AddYears(10);
            }
            else
            {
                return 1;
            }

            if (System.DateTime.Now > validTill)
            {
                return 2;
            }
            else
                return 0;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            int license = LicenseCheck();
            if (license == 1)
            {
                MessageBox.Show("Please provide valid license Key", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            else if (license == 2)
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

        private void chkWSFeed_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void btnFeed_Click(object sender, EventArgs e)
        {

        }
        
    }
}
