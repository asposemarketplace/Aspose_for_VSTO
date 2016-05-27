using Aspose.Cells;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Cells for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Cells for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            //string srcFileName = FilePath + "Conversion from docx to doc.docx";
            //string DestFileName = FilePath + "Conversion from docx to doc.doc";

            // Adding Link to a URL

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Workbook object
            int i = workbook.Worksheets.Add();

            //Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = workbook.Worksheets[i];

            //Adding a hyperlink to a URL at "A1" cell
            worksheet.Hyperlinks.Add("A1", 1, 1, "http://www.aspose.com");

            //Saving the Excel file
            workbook.Save(FilePath + "book1.xls");
        }
        static void Main2(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            // Adding a Link to a Cell in the Same File

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the first (default) worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Adding an internal hyperlink to the "B9" cell of the other worksheet "Sheet2" in
            //the same Excel file
            worksheet.Hyperlinks.Add("B3", 1, 1, "Sheet2!B9");

            //Saving the Excel file
            workbook.Save(FilePath + "book1.xls");
        }
        static void Main3(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            // Adding a Link to an External File

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Excel object
            int i = workbook.Worksheets.Add();

            //Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = workbook.Worksheets[i];

            //Adding an internal hyperlink to the "B9" cell of the other worksheet "Sheet2" in
            //the same Excel file
            worksheet.Hyperlinks.Add("A5", 1, 1, "C:\\book1.xls");

            //Saving the Excel file
            workbook.Save(FilePath + "book2.xls");
        }
    }
}
