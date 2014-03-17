using Aspose.Diagram;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
namespace Extract_All_Images_From_a_Visio
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Call the diagram constructor to load diagram from a VSD file
            Diagram diagram = new Diagram(MyDir+"ExtractImageFromShape.vsd");

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
                        bitmap.Save(MyDir+"ExtractedShape" + shape.ID + ".bmp");
                    }
                }
            }
         //   Console.ReadKey();
        }
    }
}
