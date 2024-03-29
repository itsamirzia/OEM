﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hyland.Unity;
using Hyland.Types;

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
		string basePath = string.Empty;

		//make the constructor private so that this class cannot be
		//instantiated
		private OBConnect() { }

        //Get the only object available
        public static OBConnect getInstance()
        {
            return instance;
        }

        public List<string> GetDocumentTypeGroupList(ref string obError)
        {
			List<string> documentTypeGroups = new List<string>();
			try
			{
				DocumentTypeGroupList dtgl = app.Core.DocumentTypeGroups;
				foreach (DocumentTypeGroup dtg in dtgl)
				{
					if (dtg.Name.ToUpper() == "SYSTEM DOCUMENTS")
						continue;
					documentTypeGroups.Add(dtg.ID.ToString() + " " + dtg.Name);
				}
				return documentTypeGroups;
			}
			catch (Exception ex)
			{
				obError = ex.Message;	
			}
			return documentTypeGroups;
        }
		public List<string> GetDocumentTypeList(long docTypeGroupID, ref string obError)
		{
			List<string> documentTypes = new List<string>();
			try
			{
				DocumentTypeList dtl = null;
				if (docTypeGroupID > 0)
				{
					DocumentTypeGroup dtGroup = app.Core.DocumentTypeGroups.Find(docTypeGroupID);
					dtl = dtGroup.DocumentTypes;
				}
				else
				{
					dtl = app.Core.DocumentTypes;
				}
				foreach (DocumentType dt in dtl)
				{
					if (dt.Name.Trim().ToUpper().StartsWith("SYS"))
						continue;
					documentTypes.Add(dt.ID.ToString() + " " + dt.Name);
				}

				return documentTypes;
			}
			catch (Exception ex)
			{
				obError = ex.Message;
				return documentTypes;
			}
		}
		private bool SaveToDiscWithAnnotation(Document doc, bool isAnnotationOn)
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
					string fullPath = basePath + pageData.Extension;
					Utility.WriteStreamToFile(pageData.Stream, fullPath);
				}
			}

			return true;
		}
		private bool SaveToDiscWithoutAnnotation(Document doc)
		{
			DocumentType docType = doc.DocumentType;
			if (docType.CanI(DocumentTypePrivileges.DocumentViewing))
			{
				Rendition rendition = doc.DefaultRenditionOfLatestRevision;

				DefaultDataProvider defaultDataProvider = app.Core.Retrieval.Default;

				using (PageData pageData = defaultDataProvider.GetDocument(rendition))
				{
					string fullPath = basePath + pageData.Extension;
					Utility.WriteStreamToFile(pageData.Stream, fullPath);	
				}
			}

			return true;
		}
		private DocumentList GetDocumentList(List<string> documentTypeList, DateTime from, DateTime to)
		{
			DocumentQuery docQuery = app.Core.CreateDocumentQuery();
			foreach (string docType in documentTypeList)
			{
				DocumentType dt = app.Core.DocumentTypes.Find(docType);
				docQuery.AddDocumentType(dt);
			}
			docQuery.AddDateRange(from, to);
			return docQuery.Execute(long.MaxValue);
		}
		private void ExportToNetworkLocation(DocumentList docList, bool isAnnotationOn)
		{
			foreach (Document doc in docList)
			{
				if (isAnnotationOn)
				{
					SaveToDiscWithAnnotation(doc, true);
				}
				else
				{
					SaveToDiscWithoutAnnotation(doc);
				}
			}
		}
		public bool ExportDocument(string exportPath, List<string> documentTypeList,DateTime rangeFrom, DateTime rangeTo, bool isAnnotationOn)
		{
			basePath = exportPath;
			DocumentList docList = GetDocumentList(documentTypeList, rangeFrom, rangeTo);
			ExportToNetworkLocation(docList, isAnnotationOn);
			return true;
		}

		public Hyland.Unity.Application Connect(string username, string password)
		{
			try
			{
				userName = username;
				passWord = password;
				OnBaseAuthenticationProperties authProps = Hyland.Unity.Application.CreateOnBaseAuthenticationProperties(appURL, userName, passWord, dataSource);
				app = Hyland.Unity.Application.Connect(authProps);
				return app;
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
				throw new Exception("There was an unhandled exception with the Unity API.", ex);
			}
			catch (Exception ex)
			{
				throw new Exception("There was an unhandled exception.", ex);
			}
		}

		public void Disconnect(Hyland.Unity.Application app)
		{
			if(app != null)
				app.Disconnect();
		}
	}
}
