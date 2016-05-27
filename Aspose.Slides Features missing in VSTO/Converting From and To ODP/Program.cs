using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
      private static  string FilePath = @"..\..\..\Sample Files\";
        static void Main(string[] args)
        {
            ConvertedFromOdp();
            ConvertedToOdp();
        }
        public static void  ConvertedToOdp()
        {
            string srcFileName = FilePath + "Sample Presentation.pptx";
            string destFileName = FilePath + "Output.odp";
            
            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(srcFileName))
            {

                //Saving the PPTX presentation to PPTX format
                pres.Save(destFileName, Aspose.Slides.Export.SaveFormat.Odp);
            }
        }
        public static void  ConvertedFromOdp()
        {
            string srcFileName = FilePath + "Sample Presentation.odp";
            string destFileName = FilePath + "Output.pptx";
            
            //Instantiate a Presentation object that represents a presentation file
           using(Presentation pres = new Presentation(srcFileName))
           {

               //Saving the PPTX presentation to PPTX format
              pres.Save(destFileName,Aspose.Slides.Export.SaveFormat.Pptx);
           }
        }
    }
}
