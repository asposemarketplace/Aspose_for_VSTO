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
            //Instantiate an instance of license and set the license file
            //through its path
            //Aspose.Cells.License license = new Aspose.Cells.License();
            //license.SetLicense("Aspose.Total.lic");
            //Specify the template excel file path.
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "AddWorksheetandActivate.xlsx";

            //Instantiate a new Workbook.
            //Open the excel file.
            Workbook workbook = new Workbook(fileName);

            //Declare a Worksheet object.
            Worksheet newWorksheet;

            //Add 5 new worksheets to the workbook and fill some data
            //into the cells.
            for (int i = 0; i < 5; i++)
            {

                //Add a worksheet to the workbook.
                newWorksheet = workbook.Worksheets[workbook.Worksheets.Add()];

                //Name the sheet.
                newWorksheet.Name = "New_Sheet" + (i + 1).ToString();

                //Get the Cells collection.
                Aspose.Cells.Cells cells = newWorksheet.Cells;

                //Input a string value to a cell of the sheet.
                cells[i, i].PutValue("New_Sheet" + (i + 1).ToString());
            }

            //Activate the first worksheet by default.
            workbook.Worksheets.ActiveSheetIndex = 0;

            //Save As the excel file.
            workbook.Save(fileName);
        }
    }
}
