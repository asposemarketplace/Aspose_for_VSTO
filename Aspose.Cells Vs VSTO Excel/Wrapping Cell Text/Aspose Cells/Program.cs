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
            //Create workbook
            Workbook workbook = new Workbook();

            //Access worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Place some text in cell A1 without wrapping
            Cell cellA1 = worksheet.Cells["A1"];
            cellA1.PutValue("Some Text Unwrapped");

            //Place some text in cell A5 wrapping
            Cell cellA5 = worksheet.Cells["A5"];
            cellA5.PutValue("Some Text Wrapped");
            Style style = cellA5.GetStyle();
            style.IsTextWrapped = true;
            cellA5.SetStyle(style);

            //Autofit rows
            worksheet.AutoFitRows();

            //Save the workbook
            workbook.Save(FilePath+"WrappingCellText.xlsx", SaveFormat.Xlsx);

        }
    }
}
