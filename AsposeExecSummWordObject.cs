using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Web;
using Newtonsoft.Json.Linq;

namespace SSIP_Document_Processor
{

    public class DLL
    {
        public string dll_one { get; set; }
        public string dll_two { get; set; }
        public string dll_three { get; set; }
    }
    public class SSIP_Object
    {
        public string ID { get; set; }
        public string FY { get; set; }
        public string ReleaseNumber { get; set; }
        public string VersionNumber { get; set; }
        public string BuildNumber { get; set; }
        public string Date { get; set; }
        public string DLLPath { get; set; }
        public List<DLL> DLLs { get; set; }
        public string Project { get; set; }
        public string WSPInstructions { get; set; }
        public string SVC { get; set; }
        public string WebConfig { get; set; }
        public string Database { get; set; }
        public string SSIS { get; set; }
        public string WSP { get; set; }

        public SSIP_Object(
            string ID,
            string FY,
            string ReleaseNumber,
            string VersionNumber,
            string BuildNumber,
            string Date,
            string DLLPath,
            List<DLL> DLLs,
            string Project,
            string WSPInstructions,
            string SVC,
            string WebConfig,
            string Database,
            string SSIS,
            string WSP
            )
        {
            this.ID = ID;
            this.FY = FY;
            this.ReleaseNumber = ReleaseNumber;
            this.VersionNumber = VersionNumber;
            this.BuildNumber = BuildNumber;
            this.Date = Date;
            this.DLLPath = DLLPath;
            this.DLLs = DLLs;
            this.Project = Project;
            this.WSPInstructions = WSPInstructions;
            this.SVC = SVC;
            this.WebConfig = WebConfig;
            this.Database = Database;
            this.SSIS = SSIS;
            this.WSP = WSP;
        }

        public SSIP_Object(JToken id, JToken fy, JToken releaseNumber, JToken versionNumber, JToken buildNumber, JToken date, JToken dllPath, JToken dlls, JToken project, JToken wspInstructions, JToken svc, JToken webConfig, JToken database, JToken ssis, JToken wsp)
        {
        
        }
    }

    public class AsposeExecSummWordObject
    {
        private SSIP_Object ssip_information;

        public SSIP_Object GetSsip_information()
        {
            return ssip_information;
        }

        public void SetSsip_information(SSIP_Object value)
        {
            ssip_information = value;
        }
    }
}