using System; 
using System.Collections.Generic;
using System.IO; 
using System.Data;
using System.Linq;
using System.Reflection;
using System.Web;
using Aspose.Words;
using Aspose.Words.MailMerging;



#region - SSIP_Document_Processor -

namespace SSIP_Document_Processor
{
    public class AsposeDocumentGenerator
    {

        protected string FileName { get; set; }
        protected Document SSIP_SummTemplate;
        protected License license = new License();


        public AsposeDocumentGenerator()
        {
            FileName = "";
        }

        protected void setFileName(string name)
        {
            FileName = name;
        }

        public string getFileName()
        {
            return FileName;
        }

        public Document loadWordTemplate()
        {
            Assembly a = Assembly.GetExecutingAssembly();
            Stream tempStream = a.GetManifestResourceStream(@"..\..\Content\QPP_SSIPMailMergeTemplate.docx");

            Document doc = new Document(tempStream);
            return doc;
        }

        /*public static void RegisterAspose(License license)
        {
            try
            {
                string references = "Aspose.Words.lic";
                license.SetLicense(references);
            }
            catch (Exception ex)
            {
                throw new Exception("Aspose registration failed.", ex);
            }
        }*/

        protected class HandleMergeFieldInsertXML : IFieldMergingCallback
        {
            public void FieldMerging(FieldMergingArgs args)
            {
                throw new NotImplementedException();
            }

            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                try
                {
                    if (args.DocumentFieldName.StartsWith("XML"))
                    {
                        DocumentBuilder builder = new DocumentBuilder(args.Document);
                        builder.MoveToMergeField(args.DocumentFieldName);
                        builder.InsertHtml((string)args.FieldValue, true);


                        args.Text = "";


                    }
                }
                catch (ArgumentNullException ex)
                {
                    throw new ArgumentNullException("Document field names cannot be null.", ex);
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                //Do nothing
            }
        }

        #endregion

        #region - BuildDocumentFromXML -

        public static void CreateDocumentFromXML(string configPath, string templatePath)
        {
            //Initialize DataSet
            var SSIPDocument = new DataSet();

            //Read XML into DataSet
            SSIPDocument.ReadXml(configPath);
          

            //Create the new document
            var docXML = new Document(templatePath);

            //Use each record in the table to create a new page in the document
            var customerDataTable_XML = SSIPDocument.Tables["ssip_release"]; 

            docXML.MailMerge.Execute(customerDataTable_XML);

            //Save the document
            docXML.Save(GlobalDocumentSettings.MailMergeOutputFileName_XML);
        }

        #endregion
    }
}

