using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hyland.Unity;
using System.IO;

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
		StringBuilder sbErrors = new StringBuilder();

		//make the constructor private so that this class cannot be
		//instantiated
		private OBConnect() { }

		//Get the only object available
		public static OBConnect GetInstance()
		{
			return instance;
		}
		public string RealName()
		{
			return userRealName;
		}
		//public List<string> GetDocumentTypeGroupList(ref string obError)
		//{
		//	List<string> documentTypeGroups = new List<string>();
		//	try
		//	{
		//		DocumentTypeGroupList dtgl = app.Core.DocumentTypeGroups;
		//		foreach (DocumentTypeGroup dtg in dtgl)
		//		{
		//			if (dtg.Name.ToUpper() == "SYSTEM DOCUMENTS")
		//				continue;
		//			documentTypeGroups.Add(dtg.ID.ToString() + " --- " + dtg.Name);
		//		}
		//		return documentTypeGroups;
		//	}
		//	catch (Exception ex)
		//	{
		//		obError = ex.Message;
		//	}
		//	return documentTypeGroups;
		//}
		//public List<string> GetDocumentTypeList(long docTypeGroupID, ref string obError)
		//{
		//	List<string> documentTypes = new List<string>();
		//	documentTypes.Add("0 --- All");
		//	try
		//	{
		//		DocumentTypeList dtl = null;
		//		if (docTypeGroupID > 0)
		//		{
		//			DocumentTypeGroup dtGroup = app.Core.DocumentTypeGroups.Find(docTypeGroupID);
		//			dtl = dtGroup.DocumentTypes;
		//		}
		//		else
		//		{
		//			dtl = app.Core.DocumentTypes;
		//		}
		//		foreach (DocumentType dt in dtl)
		//		{
		//			if (dt.Name.Trim().ToUpper().StartsWith("SYS"))
		//				continue;
		//			documentTypes.Add(dt.ID.ToString() + " --- " + dt.Name);
		//		}

		//		return documentTypes;
		//	}
		//	catch (Exception ex)
		//	{
		//		obError = ex.Message;
		//		return documentTypes;
		//	}
		//}
		//public bool SaveToDiscWithAnnotation(Document doc, bool isAnnotationOn)
		//{
		//	try
		//	{
		//		DocumentType docType = doc.DocumentType;
		//		if (docType.CanI(DocumentTypePrivileges.DocumentViewing))
		//		{
		//			Rendition rendition = doc.DefaultRenditionOfLatestRevision;

		//			PDFDataProvider pdfDataProvider = app.Core.Retrieval.PDF;
		//			PDFGetDocumentProperties pdfGetDocumentProperties = pdfDataProvider.CreatePDFGetDocumentProperties();
		//			pdfGetDocumentProperties.Overlay = false;
		//			pdfGetDocumentProperties.OverlayAllPages = false;
		//			pdfGetDocumentProperties.RenderNoteAnnotations = isAnnotationOn;
		//			pdfGetDocumentProperties.RenderNoteText = true;

		//			using (PageData pageData = pdfDataProvider.GetDocument(rendition, pdfGetDocumentProperties))
		//			{
		//				string fullPath = basePath + "\\" + doc.DocumentType.Name;
		//				if (!Directory.Exists(fullPath))
		//					Directory.CreateDirectory(fullPath);
		//				fullPath = fullPath + "\\" + doc.ID + "." + pageData.Extension;
		//				Utility.WriteStreamToFile(pageData.Stream, fullPath);
		//			}
		//		}

		//		return true;
		//	}
		//	catch (Exception ex)
		//	{
		//		sbErrors.AppendLine(ex.Message);
		//		return false;
		//	}
		//}
		public bool SaveToDiscWithAnnotation(string path, Document doc, bool isAnnotationOn)
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

				return true;
			}
			catch (Exception ex)
			{
				sbErrors.AppendLine(ex.Message);
				return false;
			}
		}
		//public bool SaveToDiscWithoutAnnotation(Document doc)
		//{
		//	try
		//	{
		//		DocumentType docType = doc.DocumentType;
		//		if (docType.CanI(DocumentTypePrivileges.DocumentViewing))
		//		{
		//			Rendition rendition = doc.DefaultRenditionOfLatestRevision;

		//			DefaultDataProvider defaultDataProvider = app.Core.Retrieval.Default;

		//			using (PageData pageData = defaultDataProvider.GetDocument(rendition))
		//			{
		//				string fullPath = basePath + "\\" + doc.DocumentType.Name;
		//				if (!Directory.Exists(fullPath))
		//					Directory.CreateDirectory(fullPath);
		//				string filePath = fullPath + "\\" + doc.ID + "." + pageData.Extension;
		//				Utility.WriteStreamToFile(pageData.Stream, filePath);
		//				string notes = GetNotes(doc);
		//				if (notes.Trim() != string.Empty)
		//					File.AppendAllText(fullPath + "\\" + doc.ID + ".note", notes);
		//				File.AppendAllText(fullPath + "\\" + doc.ID + ".data", GetMetaData(doc));

		//			}
		//		}

		//		return true;
		//	}
		//	catch (Exception ex)
		//	{
		//		sbErrors.AppendLine(ex.Message);
		//		return false;
		//	}
		//}
		public bool SaveToDiscWithoutAnnotation(string path, Document doc)
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
						string fullPath = path + "\\" + doc.DocumentType.Name;
						if (!Directory.Exists(fullPath))
							Directory.CreateDirectory(fullPath);
						string filePath = fullPath + "\\" + doc.ID + "." + pageData.Extension;
						Utility.WriteStreamToFile(pageData.Stream, filePath);
						string notes = GetNotes(doc);
						if (notes.Trim() != string.Empty)
							File.AppendAllText(fullPath + "\\" + doc.ID + ".note", notes);
						File.AppendAllText(fullPath + "\\" + doc.ID + ".data", GetMetaData(doc));

					}
				}

				return true;
			}
			catch (Exception ex)
			{
				sbErrors.AppendLine(ex.Message);
				return false;
			}
		}
		public List<Document> GetDocumentList(string docType, DateTime from, DateTime to, long startDH = 0, long endDH = long.MaxValue)
		{
			try
			{
				List<Document> dList = new List<Document>();
				DocumentQuery docQuery = app.Core.CreateDocumentQuery();
				if (docType.Trim().ToUpper() != "ALL")
				{
					DocumentType dt = app.Core.DocumentTypes.Find(docType);
					docQuery.AddDocumentType(dt);
				}

				docQuery.AddDateRange(from, to);
				if (startDH > 0)
					docQuery.AddDocumentRange(startDH, endDH);

				DocumentList docList = docQuery.Execute(long.MaxValue);
				foreach (Document doc in docList)
					dList.Add(doc);
				return dList;
			}
			catch
			{
				return null;
			}

		}
		//public DocumentList GetDocumentList(string docType, DateTime from, DateTime to)
		//{
		//	try
		//	{
		//		DocumentQuery docQuery = app.Core.CreateDocumentQuery();
		//		if (docType.Trim().ToUpper() != "ALL")
		//		{
		//			DocumentType dt = app.Core.DocumentTypes.Find(docType);
		//			docQuery.AddDocumentType(dt);
		//		}

		//		docQuery.AddDateRange(from, to);
		//		return docQuery.Execute(long.MaxValue);
		//	}
		//	catch
		//	{
		//		return null;
		//	}

		//}
		public string GetNotes(Document doc)
		{
			StringBuilder sbMetaData = new StringBuilder();
			try
			{
				NoteList notes = doc.Notes;
				int counter = 0;
				foreach (Note note in notes)
				{
					if (counter == 0)
						sbMetaData.AppendLine("Notes");

					string json = @"{Notes: ["
										+ "Note Title: '" + note.Title.ToString() + "',"
										+ "Created By: '" + note.CreatedBy.ToString() + "',"
										+ "Creation Date: '" + note.CreationDate.ToString() + "',"
										+ "Note Type: '" + note.NoteType.Name.ToString() + "',"
										+ "Note Page Number: '" + note.PageNumber.ToString() + "',"
										+ "Position X: '" + note.Position.X.ToString() + "',"
										+ "Position Y: '" + note.Position.Y.ToString() + "',"
										+ "Note Height: '" + note.Size.Height.ToString() + "',"
										+ "Note Width: '" + note.Size.Width.ToString() + "',"
										+ "Note Text: '" + note.Text.ToString() + "'"
									  + "]"
									+ "}";
					counter++;
				}


			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}

			return sbMetaData.ToString().Trim();
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
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}

			return metaData;
		}
		//public DocumentList GetDocumentList(long docType, DateTime from, DateTime to)
		//{
		//	DocumentQuery docQuery = app.Core.CreateDocumentQuery();
		//	if (docType != 0)
		//	{
		//		DocumentType dt = app.Core.DocumentTypes.Find(docType);
		//		docQuery.AddDocumentType(dt);
		//	}

		//	docQuery.AddDateRange(from, to);
		//	return docQuery.Execute(long.MaxValue);
		//}
		public Document GetDocumentByIDs(long DH)
		{
			try
			{
				return app.Core.GetDocumentByID(DH);
			}
			catch
			{
				return null;
			}
		}
		//public DocumentList GetDocumentList(List<string> documentTypeList, DateTime from, DateTime to)
		//{
		//	DocumentQuery docQuery = app.Core.CreateDocumentQuery();
		//	foreach (string docType in documentTypeList)
		//	{
		//		DocumentType dt = app.Core.DocumentTypes.Find(docType);
		//		docQuery.AddDocumentType(dt);
		//	}
		//	docQuery.AddDateRange(from, to);
		//	return docQuery.Execute(long.MaxValue);
		//}
		public List<Document> GetDocumentList(List<long> documentTypeList, DateTime from, DateTime to, long DHFrom, long DHTo)
		{
			List<Document> dList = new List<Document>();
			DocumentQuery docQuery = app.Core.CreateDocumentQuery();
			foreach (long docType in documentTypeList)
			{
				DocumentType dt = app.Core.DocumentTypes.Find(docType);
				docQuery.AddDocumentType(dt);
			}
			docQuery.AddDateRange(from, to);
			docQuery.AddDocumentRange(DHFrom, DHTo);
			DocumentList docList = docQuery.Execute(long.MaxValue);
			var res1 = docList.OrderBy(i => i.ID);
			var res = from dc in res1 orderby dc.ID ascending select dc;
			foreach (Document doc in res)
			{
				dList.Add(doc);
			}
			//dList = (List<Document>)dList.OrderBy(x => x.ID);
			//var res = from dc in dList orderby dc.ID ascending  select dc;
			//dList = (List<Document>)res;
			//dList = dList.Sort(x => x.ID);
			return dList;
		}
		public List<Document> GetDocumentListByDocumentTypeGroup(string DTG, DateTime from, DateTime to, long DHFrom, long DHto)
		{
			try
			{
				List<long> dtList = new List<long>();
				DocumentTypeGroup dtg = app.Core.DocumentTypeGroups.Find(DTG);
				if (dtg == null)
					throw new Exception("Document Type group not found");
				foreach (DocumentType dt in dtg.DocumentTypes)
				{
					dtList.Add(dt.ID);
				}
				return GetDocumentList(dtList, from, to, DHFrom, DHto);
			}
			catch
			{
				return null;
			}
		}
		//public void ExportToNetworkLocation(DocumentList docList, bool isAnnotationOn)
		//{
		//	foreach (Document doc in docList)
		//	{
		//		if (isAnnotationOn)
		//		{
		//			SaveToDiscWithAnnotation(doc, true);
		//		}
		//		else
		//		{
		//			SaveToDiscWithoutAnnotation(doc);
		//		}
		//	}
		//}
		//public void ExportToNetworkLocation(List<Document> docList, bool isAnnotationOn)
		//{
		//	foreach (Document doc in docList)
		//	{
		//		if (isAnnotationOn)
		//		{
		//			SaveToDiscWithAnnotation(doc, true);
		//		}
		//		else
		//		{
		//			SaveToDiscWithoutAnnotation(doc);
		//		}
		//	}
		//}
		//public bool ExportDocument(string exportPath, List<string> documentTypeList, DateTime rangeFrom, DateTime rangeTo, bool isAnnotationOn)
		//{
		//	basePath = exportPath;
		//	DocumentList docList = GetDocumentList(documentTypeList, rangeFrom, rangeTo);
		//	ExportToNetworkLocation(docList, isAnnotationOn);
		//	return true;
		//}
		//public bool ExportDocument(string exportPath, List<string> documentTypeList, string rangeFrom, string rangeTo, bool isAnnotationOn)
		//{
		//	basePath = exportPath;
		//	DateTime from = Convert.ToDateTime(rangeFrom);
		//	DateTime to = Convert.ToDateTime(rangeTo);
		//	DocumentList docList = GetDocumentList(documentTypeList, from, to);
		//	ExportToNetworkLocation(docList, isAnnotationOn);
		//	return true;
		//}
		//public bool ExportDocument(string exportPath, string documentType, string rangeFrom, string rangeTo, bool isAnnotationOn)
		//{
		//	basePath = exportPath;
		//	DateTime from = Convert.ToDateTime(rangeFrom);
		//	DateTime to = Convert.ToDateTime(rangeTo);
		//	DocumentList docList = GetDocumentList(documentType, from, to);
		//	ExportToNetworkLocation(docList, isAnnotationOn);
		//	return true;
		//}
		//public bool ExportDocument(string exportPath, List<long> documentTypeList, DateTime rangeFrom, DateTime rangeTo, bool isAnnotationOn, long DHFrom = 0, long DHTo = long.MaxValue)
		//{
		//	basePath = exportPath;
		//	List<Document> docList = GetDocumentList(documentTypeList, rangeFrom, rangeTo, DHFrom, DHTo);
		//	ExportToNetworkLocation(docList, isAnnotationOn);
		//	return true;
		//}
		//public int ExportDocument(string exportPath, string documentType, DateTime rangeFrom, DateTime rangeTo, bool isAnnotationOn)
		//{
		//	basePath = exportPath;
		//	DocumentList docList = GetDocumentList(documentType, rangeFrom, rangeTo);
		//	ExportToNetworkLocation(docList, isAnnotationOn);
		//	return docList.Count;
		//}
		//public bool ExportDocument(string exportPath, long documentType, DateTime rangeFrom, DateTime rangeTo, bool isAnnotationOn)
		//{
		//	basePath = exportPath;
		//	DocumentList docList = GetDocumentList(documentType, rangeFrom, rangeTo);
		//	ExportToNetworkLocation(docList, isAnnotationOn);
		//	return true;
		//}

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
