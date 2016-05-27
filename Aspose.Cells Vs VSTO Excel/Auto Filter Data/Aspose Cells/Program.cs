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
           //Instantiate a new Workbook.
            Workbook objBook = new Workbook();

            //Get the First sheet.
            Worksheet sheet = objBook.Worksheets["Sheet1"];

            //Add data into A1 and B1 Cells as headers.
            sheet.Cells[0, 0].PutValue("Product ID");
            sheet.Cells[0, 1].PutValue("Product Name");

            //Add data into details cells.
            sheet.Cells[1, 0].PutValue(1);
            sheet.Cells[2, 0].PutValue(2);
            sheet.Cells[3, 0].PutValue(3);
            sheet.Cells[4, 0].PutValue(4);
            sheet.Cells[1, 1].PutValue("Apples");
            sheet.Cells[2, 1].PutValue("Bananas");
            sheet.Cells[3, 1].PutValue("Grapes");
            sheet.Cells[4, 1].PutValue("Oranges");

            //Auto-filter the range.
            sheet.AutoFilter.Range = "A1:B5";

            //Auto-fit the second column.
            sheet.AutoFitColumn(1, 0, 4);

            //Save the copy of workbook as .xlsx file.
            objBook.Save(FilePath+"AutoFilterData.xlsx");
        }
    }
}
