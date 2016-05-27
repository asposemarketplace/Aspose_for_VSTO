using Aspose.Diagram;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Diagram for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Diagram for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace AsposeSourceCode.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Sample Diagram.vsdx";
            
            //Call the diagram constructor to load diagram from a VSD file
            Diagram diagram = new Diagram(srcFileName);

            //enter page index i.e. 0 for first one
            foreach (Shape shape in diagram.Pages[0].Shapes)
            {
                //Filter shapes by type Foreign
                if (shape.Type == Aspose.Diagram.TypeValue.Foreign)
                {
                    using (System.IO.MemoryStream stream = new System.IO.MemoryStream(shape.ForeignData.Value))
                    {
                        //Load memory stream into bitmap object
                        System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(stream);

                        // save bmp here
                        bitmap.Save(FilePath + "ExtractedShape" + shape.ID + ".bmp");
                    }
                }
            }
         //   Console.ReadKey();
        }
    }
}
