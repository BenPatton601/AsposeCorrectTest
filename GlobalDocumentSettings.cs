/*using Aspose.Words;
*/
namespace SSIP_Document_Processor
{
    public class GlobalDocumentSettings
    {
        //Output files
        public static string MailMergeOutputFileName_XML => @"C:\AsposeWordsDemo\SSIPMailMerge_XML.docx";
        public static string MailMergeOutputFileName_Stream => @"C:\AsposeWordsDemo\MailMergeTemplate_Stream.pdf";

        //Input files
        
        public static string MailMerge_Multiple_SSIPDocument_XML => @"..\..\Content\QPP_SSIPMailMergeTemplate.docx";
        public static string MailMergeSSIPTemplateFileName_XML => @"..\..\Content\QPP_SSIPMailMergeTemplate.docx";
        public static string CustomerXMLFilename => @"..\..\Content\SSIP.xml";
    }
}
