using Aspose.Words;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        private static string MyDir = @"..\..\..\Sample Files\";
            
        static void Main(string[] args)
        {
            ConvertingFromOdt();
            ConvertingFromOtt();
            ConvertingToOdt();
        }
        public static void ConvertingFromOdt()
        {
            Document doc = new Document(MyDir+"OpenOfficeWord.odt");
            doc.Save(MyDir+"ConvertedOdtFromDoc.docx",SaveFormat.Docx);
        }
        public static void ConvertingFromOtt()
        {
            Document doc = new Document(MyDir + "Sample.ott");
            doc.Save(MyDir + "ConvertedFromOttFromDoc.docx", SaveFormat.Docx);
        }
        public static void ConvertingToOdt()
        {
            Document doc = new Document(MyDir + "ConvertedOdtFromDoc.docx");
            doc.Save(MyDir + "ConvertedToODT.odt", SaveFormat.Odt);
        }
    }
}
