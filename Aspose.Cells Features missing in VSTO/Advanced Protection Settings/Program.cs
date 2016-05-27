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
            Workbook excel = new Workbook(fstream);

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = excel.Worksheets[0];

            //Restricting users to delete columns of the worksheet
            worksheet.Protection.AllowDeletingColumn = false;

            //Restricting users to delete row of the worksheet
            worksheet.Protection.AllowDeletingRow = false;

            //Restricting users to edit contents of the worksheet
            worksheet.Protection.AllowEditingContent = false;

            //Restricting users to edit objects of the worksheet
            worksheet.Protection.AllowEditingObject = false;

            //Restricting users to edit scenarios of the worksheet
            worksheet.Protection.AllowEditingScenario = false;

            //Restricting users to filter
            worksheet.Protection.AllowFiltering = false;

            //Allowing users to format cells of the worksheet
            worksheet.Protection.AllowFormattingCell = true;

            //Allowing users to format rows of the worksheet
            worksheet.Protection.AllowFormattingRow = true;

            //Allowing users to insert columns in the worksheet
            worksheet.Protection.AllowFormattingColumn = true;

            //Allowing users to insert hyperlinks in the worksheet
            worksheet.Protection.AllowInsertingHyperlink = true;

            //Allowing users to insert rows in the worksheet
            worksheet.Protection.AllowInsertingRow = true;

            //Allowing users to select locked cells of the worksheet
            worksheet.Protection.AllowSelectingLockedCell = true;

            //Allowing users to select unlocked cells of the worksheet
            worksheet.Protection.AllowSelectingUnlockedCell = true;

            //Allowing users to sort
            worksheet.Protection.AllowSorting = true;

            //Allowing users to use pivot tables in the worksheet
            worksheet.Protection.AllowUsingPivotTable = true;

            //Saving the modified Excel file
            excel.Save(DestFileName, SaveFormat.Xlsx);

            //Closing the file stream to free all resources
            fstream.Close();
        }
    }
}
