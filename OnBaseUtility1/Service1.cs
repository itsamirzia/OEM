using System;
using System.Collections.Generic;
using System.Data;
using System.ServiceProcess;
using Hyland.Unity;
using System.Text.RegularExpressions;

namespace OnBaseUtility1
{
    public partial class Service1 : ServiceBase
    {
        bool isConnected = false;
        Dictionary<string, string> dicDTs = new Dictionary<string, string>();
        string error = string.Empty;
        string basePath = System.Configuration.ConfigurationManager.AppSettings["BasePath"].ToString();
        bool metadataXML = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["MetadataXML"].ToString());
        string uniqueID = string.Empty;
        int docsInBatch = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DocsInBatch"].ToString());
        private System.Timers.Timer _timer;
        public Service1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Onbase Connect
        /// </summary>
        private void OBConnect()
        {
            string OBConn = System.Configuration.ConfigurationManager.AppSettings["OBConnString"].ToString();
            string[] loginArray = OBConn.Split(';');
            string appURL = string.Empty;
            string dataSource = string.Empty;
            string username = string.Empty;
            string password = string.Empty;
            bool ntAuth = false;
            bool isConnect = false;
            foreach (string str in loginArray)
            {
                string[] keyVal = str.Split('=');
                string key = keyVal[0].Trim().ToString();
                string val = keyVal[1].Trim().ToString();
                if (key.ToUpper() == "APPURL")
                    appURL = val;
                if (key.ToUpper() == "DATASOURCE")
                    dataSource = val;
                if (key.ToUpper() == "USERNAME")
                    username = val;
                if (key.ToUpper() == "PASSWORD")
                    password = val; 
                if (key.ToUpper() == "UseNTAuthentication")
                    ntAuth = Convert.ToBoolean(val);
            }
            try
            {
                OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                if (ntAuth)
                {
                    isConnect = obc.OBConnector(appURL.ToString(), dataSource.ToString(), username.ToString(), password.ToString(), true);
                }
                else
                {
                    isConnect = obc.OBConnector(appURL.ToString(), dataSource.ToString(), username.ToString(), password.ToString());
                }

                if (isConnect)
                {
                    isConnected = true;
                }
                
            }
            catch(Exception ex)
            {
                WriteToAppLogs("Exception in OnBase Connection " + ex.Message);
            }

        }
        /// <summary>
        /// Onbase Disconnect
        /// </summary>
        private void OBDisconnect()
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            obc.Disconnect();
        }
        /// <summary>
        /// Logic when service is start
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {

            if (!Environment.UserInteractive)
            {
                // Startup as service.
                _timer = new System.Timers.Timer(2 * 60 * 1000); // every 2 minute
                _timer.Elapsed += new System.Timers.ElapsedEventHandler(RunService);
                _timer.Start();
            }
            else
            {
                // Startup as application

                _timer = new System.Timers.Timer(2 * 60 * 1000); // every 2 minute
                _timer.Elapsed += new System.Timers.ElapsedEventHandler(RunService);
                _timer.Start();
            }

        }
        /// <summary>
        /// Logic after service is Start
        /// </summary>
        public void Start()
        {
            OnStart(new string[0]);
        }
        /// <summary>
        /// Run Service
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RunService(object sender, EventArgs e)
        {
            try
            {
                OBConnect();
                int licenseError = LicenseCheck();
                if (licenseError != 0)
                {
                    WriteToAppLogs("License Expired or Invalid Key");
                    this.Stop();
                }
                if (isConnected)
                {
                    MapAllData();
                    DataTable dtSearchLogs = OnBase_Export_Management.db.ExecuteSQLQuery("select * from dbo.SearchLog where Status = 'WS-Pending'");
                    foreach (DataRow dr in dtSearchLogs.Rows)
                    {
                        
                        uniqueID = dr["SearchID"].ToString().Trim();
                        OnBase_Export_Management.db.ExecuteNonQuery("update dbo.SearchLog set Status='WS-IP' where SearchID='" + uniqueID.Trim() + "'");
                        string[] docTypes = dr["DT_DTG"].ToString().Split(',');
                        string from = dr["DateRangeFrom"].ToString().Trim();
                        DateTime dtFrom = Convert.ToDateTime(from);
                        string to = dr["DateRangeTo"].ToString().Trim();
                        DateTime dtTo = Convert.ToDateTime(to);
                        long dhFrom = Convert.ToInt64(dr["StartDocHandle"].ToString().Trim());
                        long dhTo = Convert.ToInt64(dr["EndDocHandle"].ToString().Trim());
                        string isDTG = dr["isDTG"].ToString().Trim();
                        List<string> documentTypeList = new List<string>();
                        List<Document> docList = new List<Document>();
                        foreach (string documentType in docTypes)
                        {
                            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
                            WriteToAppLogs("Retrieving Document for " + documentType + ". It may take longer time. Please wait...");
                            docList.Clear();
                            docList = obc.GetDocumentList(documentType, dtFrom, dtTo, dhFrom, dhTo);
                            InertExceptionIfFound(0, uniqueID);
                            if (docList == null)
                            {

                                WriteToAppLogs("Document Type " + documentType + " not found in OnBase or user doesn't have enough rights");
                                continue;
                            }
                            if (docList.Count == 0)
                            {
                                WriteToAppLogs("No Document found for the document type " + documentType);
                                continue;
                            }

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
                                    OnBase_Export_Management.db.ExecuteNonQuery("update [dbo].[SearchLog] set end_timestamp=GETDATE(), status='WS:C' where SearchID='" + uniqueID + "';");
                                    WriteToAppLogs("Documents Downloaded successfully for Document Type " + documentType + " and date range from " + from + " to " + to);
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                OnBase_Export_Management.db.ExecuteNonQuery("Insert into dbo.exception values ('0','"+uniqueID+"','"+ex.Message+ "','WS',GETDATE())");
            }
        }
        /// <summary>
        /// License to Check
        /// </summary>
        /// <returns>Integer</returns>
        private int LicenseCheck()
        {
            DateTime releaseDate = new DateTime(2022, 12, 18);
            if (System.DateTime.Now < releaseDate)
            {
                return 1;
                //MessageBox.Show("System Date is incorrect", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //this.Close();
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
        /// <summary>
        /// Write logs to Database. Followed by WS:
        /// </summary>
        /// <param name="line"></param>
        private void WriteToAppLogs(string line)
        {
            if (Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["EnableLogsInDB"].ToString()))
                OnBase_Export_Management.db.ExecuteNonQuery("insert into [dbo].Logs values ('" + uniqueID + "','WS:" + line + "',GETDATE());");
        }
       /// <summary>
       /// Insert Exception if found.
       /// </summary>
       /// <param name="docID"></param>
       /// <param name="uID"></param>
        private void InertExceptionIfFound(long docID, string uID)
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            if (obc.CurrentException() != string.Empty)
            {
                OnBase_Export_Management.db.ExecuteNonQuery("insert into [dbo].[Exception] values (" + docID + ",'" + uniqueID + "','" + obc.CurrentException() + "','',GETDATE());");
            }

        }
        /// <summary>
        /// Convert Document Type to Dictionary
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private Dictionary<string, string> ConvertDTtoDict(DataTable dt)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (DataRow dr in dt.Rows)
            {
                dict.Add(dr[0].ToString(), dr[1].ToString());
            }
            return dict;
        }

        /// <summary>
        /// Map all data
        /// </summary>
        private void MapAllData()
        {
            OBConnector.OBConnect obc = OBConnector.OBConnect.GetInstance();
            WriteToAppLogs("Preparing Data. Please wait...");
            DataTable dtOBDT = OnBase_Export_Management.db.ExecuteSQLQuery("SELECT trim([OBDocType]),trim([ALFDocType]) FROM [dbo].[OBDocTypeVsALFDocType]");
            obc.SetOBDTvsALFDT(ConvertDTtoDict(dtOBDT));
            DataTable dtOBKey = OnBase_Export_Management.db.ExecuteSQLQuery("select trim(DocumentType)+'_'+trim(OBKey), alfkey from [dbo].[OBKeyVsALFKey]");
            obc.SetOBKeyvsALFKey(ConvertDTtoDict(dtOBKey));
            DataTable dtOBPath = OnBase_Export_Management.db.ExecuteSQLQuery("SELECT trim([OBDTG]),[DownloadPath] FROM [dbo].[OBDTGVsPath]");
            obc.SetOBDTGvsPath(ConvertDTtoDict(dtOBPath));
        }
        /// <summary>
        /// Get unique ID
        /// </summary>
        /// <returns>string</returns>
        private string GetUniqueID()
        {
            var temp = Guid.NewGuid().ToString().Replace("-", string.Empty);
            var id = Regex.Replace(temp, "[a-zA-Z]", string.Empty).Substring(0, 12);
            if (OnBase_Export_Management.db.HasDataRows("select * from [dbo].[SearchLog] where searchid='" + id + "'"))
            {
                GetUniqueID();
            }
            return id;
        }
        /// <summary>
        /// Download Document
        /// </summary>
        /// <param name="docList"></param>
        /// <param name="basePath"></param>
        /// <param name="metadataXML"></param>
        /// <returns>Boolean</returns>
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
                        string queryInsertNewRecord = @"insert into [OB_Extraction].[dbo].[SearchLog] SELECT  '" + uniqueID + "',[DT_DTG],[DateRangeFrom],[DateRangeTo],[StartDocHandle],[EndDocHandle],[IsDTG],[run_timestamp],[DocCount],[LastExecutedDH],[end_timestamp],'WS-IP',[RunByUserName] FROM [OB_Extraction].[dbo].[SearchLog] where SearchID='" + tempID + "'";
                        string queryUpdate = "update[dbo].[SearchLog] set Status = 'WS-C',end_timestamp=GETDATE(), DocCount='" + docsInBatch + "' where SearchID = '" + tempID + "'";
                        OnBase_Export_Management.db.ExecuteNonQuery(queryInsertNewRecord);
                        OnBase_Export_Management.db.ExecuteNonQuery(queryUpdate);
                        counter = 1;
                    }
                    bool downloadStatus = false;
                    if (obc.SaveToDiscWithoutAnnotation(basePath, doc, uniqueID, metadataXML))
                    {
                        downloadStatus = true;
                    }
                    if (downloadStatus)
                    {
                        WriteToAppLogs("Document Downloaded - Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);
                        if (obc.CurrentException() == string.Empty)
                        {
                            OnBase_Export_Management.db.ExecuteNonQuery("insert into [dbo].[DownloadedItems] values (" + doc.ID + ",'" + basePath + "\\" + doc.DocumentType.Name + "',GETDATE(),'WS-Success'," + uniqueID + ");");
                            OnBase_Export_Management.db.ExecuteNonQuery("update [dbo].[SearchLog] set LastExecutedDH=" + doc.ID + ", DocCount='" + counter + "' where SearchID='" + uniqueID + "';");
                        }
                        else
                        {
                            OnBase_Export_Management.db.ExecuteNonQuery("insert into [dbo].[Exception] values (" + doc.ID + ",'" + uniqueID + "','WS:" + obc.CurrentException() + "','',GETDATE());");

                        }
                    }
                    else
                    {

                        if (obc.CurrentException() != string.Empty)
                        {
                            OnBase_Export_Management.db.ExecuteNonQuery("insert into [dbo].[Exception] values (" + doc.ID + ",'" + uniqueID + "','WS:" + obc.CurrentException() + "','',GETDATE());");
                            if (obc.GetRetryCounter() >= 5)
                            {
                                string queryUpdate = "update[dbo].[SearchLog] set Status = 'WS-Aborted',end_timestamp=GETDATE() where SearchID = '" + uniqueID + "'";
                                OnBase_Export_Management.db.ExecuteNonQuery(queryUpdate);
                                return false;
                            }
                        }
                        WriteToAppLogs("Failed to Download document with Document Handle " + doc.ID + " and Document Type = " + doc.DocumentType.Name);

                        OnBase_Export_Management.db.ExecuteNonQuery("insert into [dbo].[DownloadedItems] values (" + doc.ID + ",'" + basePath + "\\" + uniqueID + "',GETDATE(),'WS-Failed');");


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
        protected override void OnStop()
        {
            OBDisconnect();
        }
    }
}
