using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
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
            //Specify the template Excel file path.
            string fileName = FilePath + "HideUnhideWorksheet.xlsx";
            //Instantiate a new Workbook.
            Workbook workbook = new Workbook(fileName);

            //Get the first sheet.
            Aspose.Cells.Worksheet objSheet = workbook.Worksheets["Sheet1"];

            //Hide the worksheet.
            objSheet.IsVisible = false;

            //Unhide the worksheet.
            objSheet.IsVisible = true;

            //Save As the Excel file.
            workbook.Save("HideUnhideWorksheet.xlsx");
        }
    }
}
