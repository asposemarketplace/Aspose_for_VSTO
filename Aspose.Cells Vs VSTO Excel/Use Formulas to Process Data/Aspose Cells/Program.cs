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
            string FilePath = @"..\..\..\..\Sample Files\";
            
            //Create workbook
            Workbook workbook = new Workbook();

            //Access worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Access cells A1, A2, A3 , A4
            Cell cellA1 = worksheet.Cells["A1"];
            Cell cellA2 = worksheet.Cells["A2"];
            Cell cellA3 = worksheet.Cells["A3"];
            Cell cellA4 = worksheet.Cells["A4"];

            //Set integer values in cells A1, A2 and A3
            cellA1.Value = 10;
            cellA2.Value = 20;
            cellA3.Value = 30;

            //Add formula in cell A4
            cellA4.Formula = "=Sum(A1:A3)";

            //Set the font bold in cell A4
            //and set the background color to Yellow in cell A4
            Style style = cellA4.GetStyle();
            style.Font.IsBold = true;
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.Yellow;
            cellA4.SetStyle(style);

            //Save the workbook
            workbook.Save(FilePath+"FormulaFunctiontoProcessData.xlsx", SaveFormat.Xlsx);
        }
    }
}
