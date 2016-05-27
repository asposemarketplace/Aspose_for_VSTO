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
            //Instantiate a new Workbook object.
            Workbook workbook = new Workbook();
            //Get the First sheet.
            Worksheet worksheet = workbook.Worksheets[0];

            //Define A1 Cell.
            Aspose.Cells.Cell cell = worksheet.Cells["A1"];
            //Add a hyperlink to it.
            int index = worksheet.Hyperlinks.Add("A1", 1, 1, "http://www.aspose.com/");
            worksheet.Hyperlinks[index].TextToDisplay = "Aspose Site!";
            worksheet.Hyperlinks[index].ScreenTip = "Click to go to Aspose site";

            //Save the excel file.
            workbook.Save(FilePath+"AddHyperlinktoCells.xls");
        }
    }
}
