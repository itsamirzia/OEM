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
				return key;
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
				return key;
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
				return "";
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
						string notes = GetNotes(doc);
						if (notes.Trim() != string.Empty)
							File.AppendAllText(fullPath + "\\" + doc.ID + ".note", notes);
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
						string notes = GetNotes(doc);
						if (notes.Trim() != string.Empty)
							File.AppendAllText(fullPath + "\\" + doc.ID + ".note", notes);
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

				sbErrors = string.Empty;
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				throw new Exception(ex.Message);
			}
			return sbMetaData.ToString().Trim();
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
					xmlWriter.WriteDocType("properties", null, "http://java.sun.com/dtd/web-app_2_3.dtd", null);

					xmlWriter.WriteStartElement("properties");
					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "type");
					xmlWriter.WriteString("inv:" + GetOBDTvsALFDT(doc.DocumentType.Name.ToString()).Trim());
					xmlWriter.WriteEndElement();

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "cm:title");
					xmlWriter.WriteString(doc.Name.ToString());
					xmlWriter.WriteEndElement();

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "inv:inv_document_id");
					xmlWriter.WriteString(doc.ID.ToString());
					xmlWriter.WriteEndElement();
					

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "inv:cl_gb_doc_date");
					xmlWriter.WriteString(doc.DocumentDate.ToString("yyyy-MM-dd"));
					xmlWriter.WriteEndElement();

					xmlWriter.WriteStartElement("entry");
					xmlWriter.WriteAttributeString("key", "inv:inv_batch_id");
					xmlWriter.WriteString(batchID);
					xmlWriter.WriteEndElement();


					foreach (KeywordRecord keywordRecord in doc.KeywordRecords)
					{
						
						foreach (Keyword keyword in keywordRecord.Keywords)
						{

							xmlWriter.WriteStartElement("entry");
							xmlWriter.WriteAttributeString("key","inv:"+ GetOBKeyvsALFKey(doc.DocumentType.Name.Trim() + "_" + keyword.KeywordType.Name.Trim()).Trim());
							xmlWriter.WriteString(keyword.Value.ToString());
							xmlWriter.WriteEndElement();
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
		public string CreateXMLwithKey(Document doc, string p)
		{
			try
			{
				XmlDocument xmlDoc = new XmlDocument();
				xmlDoc.CreateDocumentType("Properties","", "http://java.sun.com/dtd/properties.dtd",	null);
				XmlNode docNode = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
				xmlDoc.AppendChild(docNode);

				XmlNode rootNode = xmlDoc.CreateElement("Document");
				xmlDoc.AppendChild(rootNode);

				XmlNode props = xmlDoc.CreateElement("properties");
				rootNode.AppendChild(props);


				//XmlNode dh = xmlDoc.CreateElement("DocumentHandle");
				//dh.InnerText = doc.ID.ToString();
				//props.AppendChild(dh);

				XmlNode dhKey = xmlDoc.CreateElement("Property");
				dhKey.InnerText = "Document Handle";
				XmlNode dhVal = xmlDoc.CreateElement("Value");
				dhKey.InnerText = doc.ID.ToString();
				props.AppendChild(dhKey);
				props.AppendChild(dhVal);

				//XmlNode dd = xmlDoc.CreateElement("DocumentDate");
				//dd.InnerText = doc.DocumentDate.ToString();
				//props.AppendChild(dd);

				XmlNode ddKey = xmlDoc.CreateElement("Property");
				ddKey.InnerText = "Document Date";
				XmlNode ddVal = xmlDoc.CreateElement("Value");
				ddKey.InnerText = doc.DocumentDate.ToString();
				props.AppendChild(ddKey);
				props.AppendChild(ddVal);



				//XmlNode dt = xmlDoc.CreateElement("DocumentType");
				//dt.InnerText = doc.DocumentType.Name;
				//props.AppendChild(dt);

				XmlNode dtKey = xmlDoc.CreateElement("Property key= Document Type");
				dtKey.InnerText = "Document Type";
				XmlNode dtVal = xmlDoc.CreateElement("Value");
				dtKey.InnerText = doc.DocumentType.Name.ToString();
				props.AppendChild(dtKey);
				props.AppendChild(dtVal);


				//XmlNode dn = xmlDoc.CreateElement("DocumentName");
				//dn.InnerText = doc.Name;
				//props.AppendChild(dn);

				XmlNode dnKey = xmlDoc.CreateElement("Property");
				dnKey.InnerText = "Document Name";
				XmlNode dnVal = xmlDoc.CreateElement("Value");
				dnKey.InnerText = doc.Name.ToString();
				props.AppendChild(dnKey);
				props.AppendChild(dnVal);

				XmlNode Keywords = xmlDoc.CreateElement("Keywords");
				rootNode.AppendChild(Keywords);

				foreach (KeywordRecord keywordRecord in doc.KeywordRecords)
				{
					bool isKeyRec = false;
					XmlNode keyRec = null;
					if (keywordRecord.KeywordRecordType.RecordType == RecordType.MultiInstance)
					{
						keyRec = xmlDoc.CreateElement(keywordRecord.KeywordRecordType.Name.Replace(" ", ""));
						Keywords.AppendChild(keyRec);
						isKeyRec = true;
					}

					foreach (Keyword keyword in keywordRecord.Keywords)
					{
						if (isKeyRec)
						{
							XmlNode key = xmlDoc.CreateElement("Keyword");// keyword.KeywordType.Name.Replace(" ", ""));
							key.InnerText = keyword.KeywordType.Name;
							XmlNode val = xmlDoc.CreateElement("Value");
							val.InnerText =	keyword.Value.ToString();
							keyRec.AppendChild(key);
							keyRec.AppendChild(val);
						}
						else
						{
							XmlNode key = xmlDoc.CreateElement("Keyword");// keyword.KeywordType.Name.Replace(" ", ""));
							key.InnerText = keyword.KeywordType.Name;
							XmlNode val = xmlDoc.CreateElement("Value");
							val.InnerText = keyword.Value.ToString();
							Keywords.AppendChild(key);
							Keywords.AppendChild(val);

							////XmlNode key = xmlDoc.CreateElement(keyword.KeywordType.Name.Replace(" ", "").Replace("/", ""));
							////key.InnerText = keyword.Value.ToString();
							////Keywords.AppendChild(key);
						}
					}
				}
				XDocument xDoc = XDocument.Parse(xmlDoc.InnerXml);
				return xDoc.ToString();
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				return "";
			}
		}
		public string GetXMLMetadata(Document doc)
		{
			try
			{
				XmlDocument xmlDoc = new XmlDocument();
				xmlDoc.CreateDocumentType("Properties", "", "http://java.sun.com/dtd/properties.dtd", null);
				XmlNode docNode = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
				xmlDoc.AppendChild(docNode);

				XmlNode rootNode = xmlDoc.CreateElement("Document");
				xmlDoc.AppendChild(rootNode);

				XmlNode props = xmlDoc.CreateElement("properties");
				rootNode.AppendChild(props);


				XmlNode dh = xmlDoc.CreateElement("DocumentHandle");
				dh.InnerText = doc.ID.ToString();
				props.AppendChild(dh);
				XmlNode dd = xmlDoc.CreateElement("DocumentDate");
				dd.InnerText = doc.DocumentDate.ToString();
				props.AppendChild(dd);
				XmlNode dt = xmlDoc.CreateElement("DocumentType");
				dt.InnerText = doc.DocumentType.Name;
				props.AppendChild(dt);
				XmlNode dn = xmlDoc.CreateElement("DocumentName");
				dn.InnerText = doc.Name;
				props.AppendChild(dn);

				XmlNode Keywords = xmlDoc.CreateElement("Keywords");
				rootNode.AppendChild(Keywords);

				foreach (KeywordRecord keywordRecord in doc.KeywordRecords)
				{
					bool isKeyRec = false;
					XmlNode keyRec = null;
					if (keywordRecord.KeywordRecordType.RecordType == RecordType.MultiInstance)
					{
						keyRec = xmlDoc.CreateElement(keywordRecord.KeywordRecordType.Name.Replace(" ", ""));
						Keywords.AppendChild(keyRec);
						isKeyRec = true;
					}

					foreach (Keyword keyword in keywordRecord.Keywords)
					{
						if (isKeyRec)
						{
							XmlNode key = xmlDoc.CreateElement(keyword.KeywordType.Name.Replace(" ", ""));
							key.InnerText = keyword.Value.ToString();
							keyRec.AppendChild(key);
						}
						else
						{
							XmlNode key = xmlDoc.CreateElement(keyword.KeywordType.Name.Replace(" ", "").Replace("/",""));
							key.InnerText = keyword.Value.ToString();
							Keywords.AppendChild(key);
						}
					}
				}
				XDocument xDoc = XDocument.Parse(xmlDoc.InnerXml);
				return xDoc.ToString();
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
				return "";
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
				sbErrors = string.Empty;
				return app.Core.GetDocumentByID(DH);
			}
			catch(Exception ex)
			{
				sbErrors = ex.Message;
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
			try
			{
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
				sbErrors = string.Empty;
			}
			catch (Exception ex)
			{
				sbErrors = ex.Message;
			}
			return dList;
		}
		public List<Document> GetDocumentListByDocumentTypeGroup(string DTG, DateTime from, DateTime to, long DHFrom, long DHto)
		{
			try
			{
				List<long> dtList = new List<long>();
				DocumentTypeGroup dtg = app.Core.DocumentTypeGroups.Find(DTG.Trim());
				if (dtg == null)
					throw new Exception("Document Type group not found");
				foreach (DocumentType dt in dtg.DocumentTypes)
				{
					dtList.Add(dt.ID);
				}
				sbErrors = string.Empty;
				return GetDocumentList(dtList, from, to, DHFrom, DHto);
			}
			catch(Exception ex)
			{
				sbErrors = ex.Message;
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
