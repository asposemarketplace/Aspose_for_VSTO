using System;
using System.Collections.Generic;
using System.Drawing;
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
            string filePath = @"..\..\..\..\Sample Files\";
            
            //Instantiate a new Workbook.
            Workbook objBook = new Workbook();
            //Get the First sheet.
            Worksheet objSheet = objBook.Worksheets["Sheet1"];

            //Put some text into different cells (A2, A4, A6, A8).
            objSheet.Cells[1, 0].PutValue("Hair Lines");
            objSheet.Cells[3, 0].PutValue("Thin Lines");
            objSheet.Cells[5, 0].PutValue("Medium Lines");
            objSheet.Cells[7, 0].PutValue("Thick Lines");

            //Define a range object(A2).
            Aspose.Cells.Range _range;
            _range = objSheet.Cells.CreateRange("A2", "A2");
            //Set the borders with hair lines style.
            _range.SetOutlineBorders(CellBorderType.Hair, Color.Black);

            //Define a range object(A4).
            _range = objSheet.Cells.CreateRange("A4", "A4");
            //Set the borders with thin lines style.
            _range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

            //Define a range object(A6).
            _range = objSheet.Cells.CreateRange("A6", "A6");
            //Set the borders with medium lines style.
            _range.SetOutlineBorders(CellBorderType.Medium, Color.Black);

            //Define a range object(A8).
            _range = objSheet.Cells.CreateRange("A8", "A8");
            //Set the borders with thick lines style.
            _range.SetOutlineBorders(CellBorderType.Thick, Color.Black);

            //Auto-fit Column A.
            objSheet.AutoFitColumn(0);

            //Save the excel file.
            objBook.Save(filePath+"ApplyBorderstoCells.xls");
        }
    }
}
