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
            string fileName = FilePath + "MergeUnmergeCells.xlsx";
            Workbook workbook = new Workbook(fileName);

            //Get the range of cells i.e.., A1:C1.
            Aspose.Cells.Range rng1 = workbook.Worksheets[0].Cells.CreateRange("A1", "C1");

            //Merge the cells.
            rng1.Merge();

            Aspose.Cells.Cells rng = workbook.Worksheets[0].Cells;

            //UnMerge the cell.
            rng.UnMerge(0, 0, 1, 3);

            //Save the file.
            workbook.Save(fileName);
        
        }
    }
}
