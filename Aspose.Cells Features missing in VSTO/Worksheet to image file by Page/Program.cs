using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Cells for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Cells for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Sample File.xlsx";
            
            Workbook book = new Workbook(srcFileName);
            Worksheet sheet = book.Worksheets[0];
            Aspose.Cells.Rendering.ImageOrPrintOptions options = new Aspose.Cells.Rendering.ImageOrPrintOptions();
            options.HorizontalResolution = 200;
            options.VerticalResolution = 200;
            options.ImageFormat = System.Drawing.Imaging.ImageFormat.Tiff;

            //Sheet2Image By Page conversion
            SheetRender sr = new SheetRender(sheet, options);
            for (int j = 0; j < sr.PageCount; j++)
            {

                Bitmap pic = sr.ToImage(j);
                pic.Save(FilePath + sheet.Name + " Page" + (j + 1) + ".tiff");
            }

        }
    }
}
