using Aspose.Slides;
using System;
using System.Drawing;

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
            
            //Accessing the presentation
            Presentation pres = new Presentation(srcFileName);
            Image img = null;
            int slideIndex = 0;
            String ImageType = "";
            bool ifImageFound = false;
            for (int i = 0; i < pres.Slides.Count; i++)
            {
                slideIndex++;
                //Accessing the first slide
                ISlide sl = pres.Slides[i];
                System.Drawing.Imaging.ImageFormat Format = System.Drawing.Imaging.ImageFormat.Jpeg;
                for (int j = 0; j < sl.Shapes.Count; j++)
                {
                    // Accessing the shape with picture
                    IShape sh = sl.Shapes[j];

                    if (sh is AutoShape)
                    {
                        AutoShape ashp = (AutoShape)sh;
                        if (ashp.FillFormat.FillType == FillType.Picture)
                        {
                            img = ashp.FillFormat.PictureFillFormat.Picture.Image.SystemImage;
                            ImageType = ashp.FillFormat.PictureFillFormat.Picture.Image.ContentType;
                            ImageType = ImageType.Remove(0, ImageType.IndexOf("/") + 1);
                            ifImageFound = true;

                        }
                    }

                    else if (sh is PictureFrame)
                    {
                        PictureFrame pf = (PictureFrame)sh;
                        //if (pf.FillFormat.FillType == FillType.Picture)
                        {
                            img = pf.PictureFormat.Picture.Image.SystemImage;
                            ImageType = pf.PictureFormat.Picture.Image.ContentType;
                            ImageType = ImageType.Remove(0, ImageType.IndexOf("/") + 1);
                            ifImageFound = true;
                        }
                    }


                    //
                    //Setting the desired picture format
                    if (ifImageFound)
                    {
                        switch (ImageType)
                        {
                            case "jpeg":
                                Format = System.Drawing.Imaging.ImageFormat.Jpeg;
                                break;

                            case "emf":
                                Format = System.Drawing.Imaging.ImageFormat.Emf;
                                break;

                            case "bmp":
                                Format = System.Drawing.Imaging.ImageFormat.Bmp;
                                break;

                            case "png":
                                Format = System.Drawing.Imaging.ImageFormat.Png;
                                break;

                            case "wmf":
                                Format = System.Drawing.Imaging.ImageFormat.Wmf;
                                break;

                            case "gif":
                                Format = System.Drawing.Imaging.ImageFormat.Gif;
                                break;
                        }
                        //
                       
                        img.Save(FilePath+"ResultedImage"+"." + ImageType, Format);
                    }
                    ifImageFound = false;
                }
            }
        }
    }
}
