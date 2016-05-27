using Aspose.Slides;
using System.Drawing;
using System.Drawing.Imaging;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Sample Presentation with Image.pptx";
            
            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(srcFileName);

            //Accessing a slide using its slide position
            ISlide slide = pres.Slides[1];


            //Iterate all shapes on a slide and create thumbnails
            IShapeCollection shapes = slide.Shapes;
            for (int i = 0; i < shapes.Count; i++)
            {
                IShape shape = shapes[i];
                //Getting the thumbnail image of the shape
                Bitmap img = slide.GetThumbnail((float)1.0, (float)1.0);
                //Saving the thumbnail image in gif format
                img.Save(FilePath + i + ".bmp", ImageFormat.Bmp);
            }
        }
    }           
}
