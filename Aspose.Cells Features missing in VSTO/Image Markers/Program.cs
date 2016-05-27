using Aspose.Cells;
using System.Data;
using System.IO;

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
            string srcFileName = FilePath + "Sample File.xlsx";
            string DestFileName = FilePath + "Output.xlsx";
            
            //Get the image data.
            byte[] imageData = File.ReadAllBytes(FilePath + "asposeLogo.png");
            //Create a datatable.
            DataTable t = new DataTable("Table1");
            //Add a column to save pictures.
            DataColumn dc = t.Columns.Add("Picture");
            //Set its data type.
            dc.DataType = typeof(object);

            //Add a new new record to it.
            DataRow row = t.NewRow();
            row[0] = imageData;
            t.Rows.Add(row);

            //Add another record (having picture) to it.
            imageData = File.ReadAllBytes(FilePath + "Aspose.Cells.png");
            row = t.NewRow();
            row[0] = imageData;
            t.Rows.Add(row);

            //Create WorkbookDesigner object.
            WorkbookDesigner designer = new WorkbookDesigner();
            //Open the temple Excel file.
            designer.Workbook = new Workbook(srcFileName);
            //Set the datasource.
            designer.SetDataSource(t);
            //Process the markers.
            designer.Process();
            //Save the Excel file.
            designer.Workbook.Save(DestFileName);
        }
    }
}
