using System;

namespace SSIP_Document_Processor
{
    public class BuildSSIP_Document
    {
        public static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: SSIP Executable <configPath> <templatePath>" + args.Length);
            }
            else 
            {
                    string configPath = args[0];
                    string templatePath = args[1];
                    Console.WriteLine(configPath);
                    Console.WriteLine(templatePath);
                    /*SSIP_ExecSummDocGenerator test = new SSIP_ExecSummDocGenerator();
                    test.LoadJsonData();*/

                    AsposeDocumentGenerator.CreateDocumentFromXML(configPath, templatePath);
            }
            
            //Aspose document generator
            /*const string name = "dog";*/
        }
    }
}
