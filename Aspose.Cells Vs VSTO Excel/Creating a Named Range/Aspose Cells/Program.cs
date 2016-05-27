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
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.Worksheets[0];

            //Creating a named range
            Range range = worksheet.Cells.CreateRange("A1", "B4");

            //Setting the name of the named range
            range.Name = "Test_Range";

            for (int row = 0; row < range.RowCount; row++)
            {
                for (int column = 0; column < range.ColumnCount; column++)
                {
                    range[row, column].PutValue("Test");
                }
            }

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.Save(FilePath+"CreateNamedRange.xls");
        }
    }
}
