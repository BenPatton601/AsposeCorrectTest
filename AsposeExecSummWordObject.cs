using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Web; 

namespace SSIP_Document_Processor
{
    public class AsposeExecSummWordObject
    {
        public string ID { get; set; }
        public string FY { get; set; }
        public string ReleaseNumber { get; set; }
        public string VersionNumber { get; set; }
        public string BuildNumber { get; set; }
        public string Date { get; set; }
        public string DLLPath { get; set; }
        public Array DLLs { get; set; }
        public string Project { get; set; }
        public string WSPInstructions { get; set; }
        public string SVC { get; set; }
        public string WebConfig { get; set; }
        public string Database { get; set; }
        public string SSIS { get; set; }
        public string WSP { get; set; }
        public AsposeExecSummWordObject()
        {
             
        }

    }
}
