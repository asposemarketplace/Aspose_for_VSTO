using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            string fileName = FilePath + "ApplyStyletoRanges.xlsx";

            Workbook myWorkbook = new Workbook(fileName);
            Worksheet mySheet = myWorkbook.Worksheets[myWorkbook.Worksheets.ActiveSheetIndex];

            Style style = myWorkbook.CreateStyle();
            style.VerticalAlignment = TextAlignmentType.Center;
            //Setting the horizontal alignment of the text in the "A1" cell
            style.HorizontalAlignment = TextAlignmentType.Center;
            //Setting the font color of the text in the "A1" cell
            style.Font.Color = Color.Green;
            //Shrinking the text to fit in the cell
            style.ShrinkToFit = true;
            //Setting the bottom border color of the cell to red
            style.Borders[BorderType.BottomBorder].Color = Color.Red;

            //Creating StyleFlag
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.HorizontalAlignment = true;
            styleFlag.VerticalAlignment = true;
            styleFlag.ShrinkToFit = true;
            styleFlag.Borders = true;
            styleFlag.FontColor = true;

            //Accessing a row from the Rows collection
            Column column = mySheet.Cells.Columns[0];
            //Assigning the Style object to the Style property of the row
            column.ApplyStyle(style, styleFlag);

            myWorkbook.Save(fileName);
        }
    }
}
