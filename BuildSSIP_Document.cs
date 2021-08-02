using System;

namespace SSIP_Document_Processor
{
    public class BuildSSIP_Document
    {
        public static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: SSIP Executable <configPath> <templatePath>" + args.Length);
            }
            else 
            {

                SSIP_ExecSummDocGenerator ssip_document = new SSIP_ExecSummDocGenerator(); 

                    string configPath = args[0];
                    string templatePath = args[1];
                    string jsonConfigPath = args[2]; 
                    Console.WriteLine(configPath);
                    Console.WriteLine(templatePath);
                    Console.WriteLine(jsonConfigPath);  
                    /*SSIP_ExecSummDocGenerator test = new SSIP_ExecSummDocGenerator();
                    test.LoadJsonData();*/

                    /*AsposeDocumentGenerator.CreateDocumentFromXML(configPath, templatePath);*/
                    ssip_document.GenerateSSIPDocument(jsonConfigPath, templatePath); 
            }
            
            //Aspose document generator
            /*const string name = "dog";*/
        }
    }
}
