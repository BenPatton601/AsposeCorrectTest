using System; 
using System.Collections.Generic;
using System.IO; 
using System.Data;
using System.Linq;
using System.Reflection;
using System.Web;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;



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

        public Document loadWordTemplate(string templatePath)
        {
            /*Assembly a = Assembly.GetExecutingAssembly();
            Stream tempStream = a.GetManifestResourceStream(templatePath);*/

            Document doc = new Document(templatePath);
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


    /*public Object LoadJsonData(string jsonConfigPath)
    {

            *//* var test;*//*
            Object obj;

        using (StreamReader r = new StreamReader(jsonConfigPath))
        {
            string json = r.ReadToEnd();
            var test = JsonConvert.DeserializeObject(json);
            Console.WriteLine("This is the test", test);
            var jsonobject = JObject.FromObject(test);
            Console.WriteLine(jsonobject);
            Console.WriteLine(jsonobject["release_number"]);



            var id = "0";
            var fy = "2021";
            var releaseNumber = jsonobject["release_number"];
            var versionNumber = jsonobject["version_number"];
            var buildNumber = jsonobject["build_number"];
            var date = jsonobject["date"];
            var dllPath = jsonobject["dll_path"];
            var dlls = jsonobject["dlls"];
            var project = jsonobject["project"];
            var wspInstructions = jsonobject["wsp_instructions"];
            var svc = jsonobject["svc"];
            var webConfig = jsonobject["web_config"];
            var database = jsonobject["database"];
            var ssis = jsonobject["ssis"];
            var wsp = jsonobject["wsp"];

            if (releaseNumber == null)
            {
                releaseNumber = "";
            }
            if (versionNumber == null)
            {
                versionNumber = "";
            }
            if (buildNumber == null)
            {
                buildNumber = "";
            }
            if (date == null)
            {
                date = "";
            }
            if (dllPath == null)
            {
                dllPath = "";
            }
            if (dlls == null)
            {
                dlls = "";
            }
            if (project == null)
            {
                project = "";
            }
            if (wspInstructions == null)
            {
                wspInstructions = "";
            }
            if (svc == null)
            {
                svc = "";
            }
            if (webConfig == null)
            {
                webConfig = "";
            }
            if (database == null)
            {
                database = "";
            }
            if (ssis == null)
            {
                ssis = "";
            }
            if (wsp == null)
            {
                wsp = "";
            }


            obj = new Object(id, fy, releaseNumber, versionNumber, buildNumber, date, dllPath, dlls, project, wspInstructions, svc, webConfig, database, ssis, wsp);
            Console.WriteLine(obj);
        }

        return obj;
    }*/

    #region 



  /*  public Document CreateDocumentFromJSON(string templatePath, string jsonConfigPath)
    {
        //initialize Dataset
        var SSIP_JSON_Doc = new DataSet();

            //Read the JSON
            Object jsonFile = LoadJsonData(jsonConfigPath); 

    }*/





    #endregion
}
}

