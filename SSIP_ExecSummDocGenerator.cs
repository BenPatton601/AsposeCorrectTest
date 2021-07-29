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

namespace SSIP_Document_Processor
{
    public class SSIP_ExecSummDocGenerator : AsposeDocumentGenerator
    {
        //Global variables needed for execution
       /* private string[] SSIPMergeFields;
        private object[] SSIPDataColumns; */


        //Initializes the object with Facility Component Service and updates Merge Fields
       /* public SSIP_DocGenerator()
        {
            SSIPMergeFields = new string[] { "XML_ReleaseNumber", "XML_VersionNumber", "XML_Date", "XML_DLLPath", "XML_DLLs", "XML_Project", "XML_WSPInstructions", "XML_SVC", "XML_WebConfig", "XML_Database", "XML_SSIS", "XML_WSP" };
        }*/

        

        public AsposeExecSummWordObject LoadJsonData()
        {
            AsposeExecSummWordObject obj = new AsposeExecSummWordObject();
           /* var test;*/ 


            using (StreamReader r = new StreamReader("SSIP.json"))
            {
                string json = r.ReadToEnd();
                var test = (AsposeExecSummWordObject)JsonConvert.DeserializeObject(json);
                Console.WriteLine("This is the test data", test); 
            }

            /*string json = ("SSIP.json");*/
            /*obj = (AsposeExecSummWordObject)JsonConvert.DeserializeObject(json);
            Console.WriteLine(JsonConvert.DeserializeObject(json)); */

            return obj;
        }

       /* public ExecutiveSummaryDTO loadExecData(int exSummId)
        {
            ExecutiveSummaryDTO execData = new ExecutiveSummaryDTO();
            execData = facilityComponentService.GetExecutiveSummaryDTOHelper(exSummId);
            return execData;
        }

        public string[] ValidateData(string[] data)
        {
            if (data = null)
            {
                //Exit the function here
            }
            else
            {
                if (data.ReleaseNumber == null)
                {
                    data.ReleaseNumber = ""; 
                }
                if (data.VersionNumber == null)
                {
                    data.VersionNumber = ""; 
                }
                if (data.Date == null)
                {
                    data.Date = ""; 
                }
                if (data.DLLPath == null)
                {
                    data.DLLPath = ""; 
                }
                if (data.DLLs == null)
                {
                    data.DLL = ""; 
                }
                if (data.DLLs)
                {
                    string[] DLL_List = data.DLLs.Split(","); 
                    foreach (string DLL in DLL_List)
                }
            }
        }*/
    }
}
