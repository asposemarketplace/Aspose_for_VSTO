using Aspose.Cells;
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
            
            //Creating a file stream containing the Excel file to be opened
            FileStream fstream = new FileStream(srcFileName, FileMode.Open);

            //Instantiating a Workbook object
            //Opening the Excel file through the file stream
            Workbook workbook = new Workbook(fstream);

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.Worksheets[0];

            //Hiding the gridlines of the first worksheet of the Excel file
            worksheet.IsGridlinesVisible = false;

            //Saving the modified Excel file
            workbook.Save(DestFileName);

            //Closing the file stream to free all resources
            fstream.Close();
        }
    }
}
