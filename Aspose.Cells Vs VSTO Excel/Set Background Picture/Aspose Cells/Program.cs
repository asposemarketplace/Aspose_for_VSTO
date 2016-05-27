using System;
using System.Collections.Generic;
using System.IO;
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
            Workbook workbook = new Workbook();
            //Get the first worksheet. 
            Worksheet sheet = workbook.Worksheets[0];

            //Define a string variable to store the image path.
            string ImageUrl = FilePath+"image.jpg";
            //Get the picture into the streams.
            FileStream fs = File.OpenRead(ImageUrl);
            //Define a byte array.
            byte[] imageData = new Byte[fs.Length];
            //Obtain the picture into the array of bytes from streams.
            fs.Read(imageData, 0, imageData.Length);
            //Close the stream.
            fs.Close();

            //Set the background image for the sheet.
            sheet.SetBackground(imageData);

            //Save the excel file.
            workbook.Save(FilePath+"Setbackgroundpic.xlsx");
        }
    }
}
