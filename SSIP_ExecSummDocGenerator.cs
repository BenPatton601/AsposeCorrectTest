using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using System.IO;
using System.Web;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SSIP_Document_Processor
{
    public class SSIP_ExecSummDocGenerator : AsposeDocumentGenerator
    {
        //Global variables needed for execution
        private string[] SSIPMergeFields;
        private object[] SSIPDataColumns;


        //Initializes the object with Facility Component Service and updates Merge Fields
        public SSIP_ExecSummDocGenerator()
        {
            SSIPMergeFields = new string[] {
                "ID",
                "FY", 
                "ReleaseNumber",
                "VersionNumber",
                "BuildNumber",
                "Date",
                "DLLPath",
                "DLLs",
                "Project",
                "WSPInstructions",
                "SVC",
                "WebConfig",
                "Database",
                "SSIS",
                "WSP"
            };
        }



        public SSIP_Object LoadJsonData(string jsonConfigPath)
        {

            /* var test;*/
            SSIP_Object obj;

            using (StreamReader r = new StreamReader(jsonConfigPath))
            {
                Console.WriteLine(jsonConfigPath); 
                string json = r.ReadToEnd();
                var test = JsonConvert.DeserializeObject(json);
                Console.WriteLine("This is the test", test);
                var jsonobject = JObject.FromObject(test);
                Console.WriteLine(jsonobject);
                Console.WriteLine(jsonobject["release_number"]);



                var id = jsonobject["ID"];
                var fy = jsonobject["FY"];
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


                obj = new SSIP_Object(id, fy, releaseNumber, versionNumber, buildNumber, date, dllPath, dlls, project, wspInstructions, svc, webConfig, database, ssis, wsp);
                Console.WriteLine(obj);
            }

            return obj;
        }



        /*public SSIP_Object ValidateData(SSIP_Object obj)
        {
            
                if (obj.ReleaseNumber == null)
                {
                    obj.ReleaseNumber = "";
                }
                if (obj.VersionNumber == null)
                {
                    obj.VersionNumber = "";
                }
                if (obj.Date == null)
                {
                    obj.Date = "";
                }
                if (obj.DLLPath == null)
                {
                    obj.DLLPath = "";
                }
                if (obj.dlls == null)
                {
                    obj.DLLs = "";
                }
            return obj;
        }*/

        //Method sets up the data columns to merge the matching fields together with appropriate data.
        //Params:
        //data - FacilityDTO used to name the file and title the document
        //execData - ExecutiveSummaryDTO used to populate the document with data
        //FY - used to identify what year in the document to remove confusion.
        public object[] SetUpDataColumns(string jsonConfigPath)
        {
            SSIP_Object ssipData = LoadJsonData(jsonConfigPath);
            object[] SSIPDataColumns = new object[] { 
                                   ssipData.ID, ssipData.FY,
                                   ssipData.Project, ssipData.BuildNumber,
                                   ssipData.VersionNumber, ssipData.ReleaseNumber,
                                   ssipData.Date, ssipData.DLLPath,
                                   ssipData.DLLs, ssipData.WSPInstructions,
                                   ssipData.SVC, ssipData.WebConfig, 
                                   ssipData.Database,
                                   ssipData.SSIS, ssipData.WSP
            };
            return SSIPDataColumns;
        }

        //Method to merge the data compiled into the document.
        public Document MergeSSIPDataToTemplate(Document template, string jsonConfigPath)
        {
            // the DocumentBuilder class provides members to easily add content to a document
            DocumentBuilder builder = new DocumentBuilder(template);
            SSIPDataColumns = SetUpDataColumns(jsonConfigPath);
            template.MailMerge.FieldMergingCallback = new HandleMergeFieldInsertXML();
            template.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            builder.Document.MailMerge.Execute(SSIPMergeFields, SSIPDataColumns);
            return template;
        }

        //If Aspose Word Document feature is wanted for Published versions only: Clean up can occur and data calls can be minimized by grabbing FacilityComponentDTO/MarketPlanDTO
        //public Document GenerateFacilityDocument(string facilityId, string FY, string facilityComponentVersionId, string facilityComponentId)
        public Document GenerateSSIPDocument(string jsonConfigPath, string templatePath)
        {
            Console.WriteLine("testing SSIP");
            Console.WriteLine($"This is the json config path{jsonConfigPath}");
            Console.WriteLine($"This is the template path {templatePath}");
            /*            RegisterAspose(license);
            */
            SSIP_Object ssipData = LoadJsonData(jsonConfigPath);
            setFileName(ssipData.Project);
            Console.WriteLine($"This is the ssipData project name{ssipData.Project}");
            Document template = loadWordTemplate(templatePath);
            string dir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            dir += @"\";
            //Merge the necessary data into the template.
            Document doc = MergeSSIPDataToTemplate(template, jsonConfigPath);
            doc.Save(dir + ssipData.Project + ".docx");
            return doc;
        }
    }
}