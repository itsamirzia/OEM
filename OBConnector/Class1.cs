using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hyland.Unity;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace OBConnector
{
	public class OBConnect
	{

		private static OBConnect instance = new OBConnect();
		Hyland.Unity.Application app = null;
		string appURL = string.Empty;
		string dataSource = string.Empty;
		string userName = string.Empty;
		string passWord = string.Empty;
		string userRealName = string.Empty;
		string basePath = string.Empty;
		static string session = string.Empty;
		string sbErrors = string.Empty;
		Dictionary<string, string> OBDTvsALFDT = new Dictionary<string, string>();
		Dictionary<string, string> OBKeyvsALFKey = new Dictionary<string, string>();
		Dictionary<string, string> OBDTGvsPath = new Dictionary<string, string>();
		List<string> CTXDocList = new List<string>();

		//make the constructor private so that this class cannot be
		//instantiated
		private OBConnect() { }

		//Get the only object available
		public static OBConnect GetInstance()
		{
			return instance;
		}
		public void SetCTXDocList(List<string> docList)
		{
			foreach (string doc in docList)
			{
				CTXDocList.Add(doc.Trim());
			}
		}
		public void SetOBDTvsALFDT(Dictionary<string, string> dict)
		{
			OBDTvsALFDT = dict.ToDictionary(entry => entry.Key, entry => entry.Value);
		}
		public void SetOBKeyvsALFKey(Dictionary<string, string> dict)
		{
			OBKeyvsALFKey = dict.ToDictionary(entry => entry.Key, entry => entry.Value);
		}
		public void SetOBDTGvsPath(Dictionary<string, string> dict)
		{
			OBDTGvsPath = dict.ToDictionary(entry => entry.Key, entry => entry.Value);
		}
		private string GetOBDTvsALFDT(string key)
		{
			try
			{
				return OBDTvsALFDT[key];
			}
			catch
			{
				return string.Empty;
			}
		}
		private string GetOBKeyvsALFKey(string key)
		{
			try
			{
				return OBKeyvsALFKey[key];
			}
			catch
			{
				return string.Empty;
			}
		}
		private string GetOBDTGvsPath(string key)
		{
			try
			{
				return OBDTGvsPath[key];
			}
			catch
			{
				return string.Empty;
			}
		}
		public string CurrentException()
        {
			return sbErrors;
        }
		public string RealName()
		{
			return userRealName;
		}
		public bool SaveToDiscWithAnnotation(string path, Document doc, bool isAnnotationOn, bool metadataXML = true)
		{
			try
			{
				DocumentType docType = doc.DocumentType;
				if (docType.CanI(DocumentTypePrivileges.DocumentViewing))
				{
					Rendition rendition = doc.DefaultRenditionOfLatestRevision;

					PDFDataProvider pdfDataProvider = app.Core.Retrieval.PDF;
					PDFGetDocumentProperties pdfGetDocumentProperties = pdfDataProvider.CreatePDFGetDocumentProperties();
					pdfGetDocumentProperties.Overlay = false;
					pdfGetDocumentProperties.OverlayAllPages = false;
					pdfGetDocumentProperties.RenderNoteAnnotations = isAnnotationOn;
					pdfGetDocumentProperties.RenderNoteText = true;

					using (PageData pageData = pdfDataProvider.GetDocument(rendition, pdfGetDocumentProperties))
					{
						string fullPath = path + "\\" + doc.DocumentType.Name;
						if (!Directory.Exists(fullPath))
							Directory.CreateDirectory(fullPath);
						fullPath = fullPath + "\\" + doc.ID + "." + pageData.Extension;
						Utility.WriteStreamToFile(pageData.Stream, fullPath);
					}
				}
				sbErrors = string.Empty;
				return true;
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				return false;
			}
		}
		public bool SaveTXTToDiscWithoutAnnotation(string path, Document doc, string batchid, bool metadataXML = true)
		{

			try
			{
				DocumentType docType = doc.DocumentType;
				if (docType.CanI(DocumentTypePrivileges.DocumentViewing))
				{
					
					//Revision rev = doc.Revisions.ElementAt(0);
					//Rendition rend = rev.DefaultRendition;
					Rendition rendition = doc.DefaultRenditionOfLatestRevision;

					TextDataProvider defaultDataProvider = app.Core.Retrieval.Text;

					using (PageData pageData = defaultDataProvider.GetDocument(rendition))
					{
						string fullPath = path + "\\" + batchid + "\\" + GetOBDTGvsPath(doc.DocumentType.Name);
						if (!Directory.Exists(fullPath))
							Directory.CreateDirectory(fullPath);
						string filePath = fullPath + "\\" + doc.ID + ".txt";
						Utility.WriteStreamToFile(pageData.Stream, filePath);
						//string notes = GetNotes(doc);
						//if (notes.Trim() != string.Empty)
						//	File.AppendAllText(fullPath + "\\" + doc.ID + ".note", notes);
						if (metadataXML)
						{
							if (!CreateXMLwithKey(doc, fullPath + "\\" + doc.ID+"."+pageData.Extension.ToLower() + ".metadata.properties.xml", batchid))
							{
								sbErrors += " \r\n Document Downloaded with but issue with XML";
							}
							else
								sbErrors = string.Empty;
						}
						else
						{
							File.AppendAllText(fullPath + "\\" + doc.ID + ".Metadata", GetMetaData(doc));
						}

					}
				}

				return true;
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				return false;
			}
		}
		public bool SaveToDiscWithoutAnnotation(string path, Document doc, string batchid, bool metadataXML = true)
		{
			try
			{
				DocumentType docType = doc.DocumentType;
				if (docType.CanI(DocumentTypePrivileges.DocumentViewing))
				{
					Rendition rendition = doc.DefaultRenditionOfLatestRevision;

					DefaultDataProvider defaultDataProvider = app.Core.Retrieval.Default;

					using (PageData pageData = defaultDataProvider.GetDocument(rendition))
					{
						string fullPath = path + "\\" + batchid + "\\"+GetOBDTGvsPath(doc.DocumentType.Name);
						if (!Directory.Exists(fullPath))
							Directory.CreateDirectory(fullPath);
						string filePath = fullPath + "\\" + doc.ID + "." + pageData.Extension;
						Utility.WriteStreamToFile(pageData.Stream, filePath);
						//string notes = GetNotes(doc);
						//if (notes.Trim() != string.Empty)
						//	File.AppendAllText(fullPath + "\\" + doc.ID + ".note", notes);
						if (metadataXML)
						{
							if (!CreateXMLwithKey(doc, fullPath + "\\" + doc.ID +"."+pageData.Extension+ ".metadata.properties.xml", batchid))
							{
								sbErrors += " \r\n Document Downloaded with but issue with XML";
							}
							else
								sbErrors = string.Empty;
						}
						else
						{
							File.AppendAllText(fullPath + "\\" + doc.ID + ".Metadata", GetMetaData(doc));
						}

                    }
				}
				
				return true;
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				return false;
			}
		}
		private long GetDocumentTypeNumber(string number)
		{
			try
			{
				return Convert.ToInt64(number.Trim());
			}
			catch
			{
				return 0;
			}
		}
		public List<Document> GetDocumentList(string docType, DateTime from, DateTime to, long startDH = 0, long endDH = long.MaxValue)
		{
			bool DTisNumber = false;
			long DTNumber = GetDocumentTypeNumber(docType);
			if (DTNumber != 0)
				DTisNumber = true;
			List<Document> dList = new List<Document>();
			try
			{
				

				DocumentQuery docQuery = app.Core.CreateDocumentQuery();
				if (docType.Trim().ToUpper() != "ALL")
				{
					DocumentType dt = null;
					if (DTisNumber)
						dt = app.Core.DocumentTypes.Find(DTNumber);
					else
						dt = app.Core.DocumentTypes.Find(docType.Trim());

					if (dt == null)
						throw new Exception(docType + " not found in OnBase.");
					docQuery.AddDocumentType(dt);
				}

				docQuery.AddDateRange(from, to);
				if (startDH > 0)
					docQuery.AddDocumentRange(startDH, endDH);

				DocumentList docList = docQuery.Execute(long.MaxValue);
				foreach (Document doc in docList)
					dList.Add(doc);
				sbErrors = string.Empty;
				return dList;
			}
			catch(Exception ex)
			{
				sbErrors = ex.Message;
				return dList;
			}

		}		
		public bool CreateXMLwithKey(Document doc, string fullpath, string batchID)
		{
			StringBuilder sbXML = new StringBuilder();
			try
			{
				XmlWriterSettings settings = new XmlWriterSettings();
				settings.ConformanceLevel = ConformanceLevel.Document;
				settings.Indent = true;
				settings.IndentChars = "\t";
				

				using (XmlWriter xmlWriter = XmlWriter.Create(fullpath, settings))
				{

					xmlWriter.WriteStartDocument();
					xmlWriter.WriteDocType("properties", null, "http://java.sun.com/dtd/properties.dtd", null);

					xmlWriter.WriteStartElement("properties");
					xmlWriter.WriteStartElement("entry");
					string alfDocType = GetOBDTvsALFDT(doc.DocumentType.Name.ToString()).Trim();
					if (alfDocType != string.Empty)
					{
						xmlWriter.WriteAttributeString("key", "type");
						xmlWriter.WriteString("inv:" + alfDocType);
						xmlWriter.WriteEndElement();
					}

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "cm:title");
					xmlWriter.WriteString(doc.Name.ToString());
					xmlWriter.WriteEndElement();

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "inv:inv_document_id");
					xmlWriter.WriteString(doc.ID.ToString());
					xmlWriter.WriteEndElement();
					

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "inv:inv_creation_date");
					xmlWriter.WriteString(doc.DocumentDate.ToString("yyyy-MM-dd"));
					xmlWriter.WriteEndElement();

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "inv:inv_batch_id");
					xmlWriter.WriteString(batchID);
					xmlWriter.WriteEndElement();

					string obAlfDocType = GetOBKeyvsALFKey(doc.DocumentType.Name.Trim() + "_Document Type Name");
					if (obAlfDocType != string.Empty)
					{
						xmlWriter.WriteStartElement("entry");
						xmlWriter.WriteAttributeString("key", "inv:"+obAlfDocType.Trim().ToLower());
						xmlWriter.WriteString(doc.DocumentType.Name);
						xmlWriter.WriteEndElement();
					}
					string obAlfDocTypeGroup = GetOBKeyvsALFKey(doc.DocumentType.Name.Trim() + "_Document Group Name");
					if (obAlfDocTypeGroup != string.Empty)
					{
						xmlWriter.WriteStartElement("entry");
						xmlWriter.WriteAttributeString("key", "inv:" + obAlfDocTypeGroup.Trim().ToLower());
						xmlWriter.WriteString(doc.DocumentType.DocumentTypeGroup.Name);
						xmlWriter.WriteEndElement();
					}


					foreach (KeywordRecord keywordRecord in doc.KeywordRecords)
					{
						
						foreach (Keyword keyword in keywordRecord.Keywords)
						{
							string keywordValue = string.Empty;
							if ((keyword.KeywordType.DataType == KeywordDataType.DateTime) || (keyword.KeywordType.DataType == KeywordDataType.Date))
							{
								keywordValue = Convert.ToDateTime(keyword.Value).ToString("yyyy-MM-dd");
							}
							else
								keywordValue = keyword.Value.ToString();

							string mapKeyword = GetOBKeyvsALFKey(doc.DocumentType.Name.Trim() + "_" + keyword.KeywordType.Name.Trim());
							if (mapKeyword != string.Empty)
							{
								xmlWriter.WriteStartElement("entry");
								xmlWriter.WriteAttributeString("key", "inv:" + mapKeyword.Trim().ToLower());
								xmlWriter.WriteString(keywordValue);
								xmlWriter.WriteEndElement();
							}
						}
					}

						//xmlWriter.WriteEndElement();

					xmlWriter.WriteEndElement();
					xmlWriter.WriteEndDocument();
				}
				return true;

			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				return false;
			}
		}
		public string GetMetaData(Document doc)
		{
			string metaData = string.Empty;
			StringBuilder sbMetaData = new StringBuilder();
			sbMetaData.AppendLine("{");
			sbMetaData.AppendLine("Document Handle: " + doc.ID);
			sbMetaData.AppendLine("Document Date: " + doc.DocumentDate);
			sbMetaData.AppendLine("Document Type: " + doc.DocumentType.Name);
			sbMetaData.AppendLine(@"Keywords:[");
			try
			{
				foreach (KeywordRecord keywordRecord in doc.KeywordRecords)
				{
					foreach (Keyword keyword in keywordRecord.Keywords)
					{
						string sKeyValue = keyword.IsBlank ? string.Empty : keyword.Value.ToString();
						sbMetaData.AppendLine(keyword.KeywordType.Name + " : '" + sKeyValue + "',");
					}
				}
				metaData = sbMetaData.ToString().TrimEnd(',') + "]}";
				sbErrors = string.Empty;
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				throw new Exception(ex.Message);
			}

			return metaData;
		}
		public Document GetDocumentByIDs(long DH)
		{
			try
			{
				sbErrors = string.Empty;
				return app.Core.GetDocumentByID(DH);
			}
			catch(Exception ex)
			{
				sbErrors = ex.Message;
				return null;
			}
		}
		public bool Connect(string appUrl, string DataSource, string username, string password, bool NTAuth = false)
		{
			try
			{
				appURL = appUrl;
				dataSource = DataSource;
				userName = username;
				passWord = password;

				if (NTAuth)
				{
					DomainAuthenticationProperties dap = Hyland.Unity.Application.CreateDomainAuthenticationProperties(appUrl, dataSource);
					app = Hyland.Unity.Application.Connect(dap);
				}
				else
				{
					OnBaseAuthenticationProperties authProps = Hyland.Unity.Application.CreateOnBaseAuthenticationProperties(appURL, userName, passWord, dataSource);
					app = Hyland.Unity.Application.Connect(authProps);
					
				}
				userRealName = app.CurrentUser.RealName;
				session = app.SessionID;
				return true;
			}
			catch (InvalidLoginException ex)
			{
				throw new Exception("The credentials entered are invalid.", ex);
			}
			catch (AuthenticationFailedException ex)
			{
				throw new Exception("Authentication failed.", ex);
			}
			catch (MaxConcurrentLicensesException ex)
			{
				throw new Exception("All licenses are currently in use, please try again later.", ex);
			}
			catch (NamedLicenseNotAvailableException ex)
			{
				throw new Exception("Your license is not availble, please insure you are logged out of other OnBase clients.", ex);
			}
			catch (SystemLockedOutException ex)
			{
				throw new Exception("The system is currently locked, please try back later.", ex);
			}
			catch (UnityAPIException ex)
			{
				throw new Exception(ex.Message);
			}
			catch (Exception ex)
			{
				throw new Exception("Default exception." + ex.Message, ex);
			}
		}

		public void Disconnect()
		{
			if (app != null)
			{
				app.Disconnect();
				app = null;
			}
		}
	}
}
