using Aspose.Words;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Create the headers.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header Text goes here...");
            //add footer having current date
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.InsertField("Date", "");

            doc.UpdateFields();
            doc.Save(FilePath + "Insert Headers and Footers.docx");
        }
    }
}
