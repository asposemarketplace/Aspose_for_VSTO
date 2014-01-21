using Aspose.Diagram;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Diagram
{
    class Program
    {
        static void Main(string[] args)
        {
            string visioStencil = "sample.vss";
            // Create a new diagram
            Diagram diagram = new Diagram(visioStencil);

            //Add a new rectangle shape
            long shapeId = diagram.AddShape(
                4.25, 5.5, 2, 1, @"Rectangle", 0);
            Shape shape = diagram.Pages[0].Shapes.GetShape(shapeId);
            shape.Text.Value.Add(new Txt(@"Rectangle text."));

            //Add a new star shape
            shapeId = diagram.AddShape(
                2.0, 5.5, 2, 2, @"Star 7", 0);
            shape = diagram.Pages[0].Shapes.GetShape(shapeId);
            shape.Text.Value.Add(new Txt(@"Star text."));

            //Add a new hexagon shape
            shapeId = diagram.AddShape(
              7.0, 5.5, 2, 2, @"Hexagon", 0);
            shape = diagram.Pages[0].Shapes.GetShape(shapeId);
            shape.Text.Value.Add(new Txt(@"Hexagon text."));

            //Save the new diagram
            diagram.Save("Drawing1.vdx", SaveFileFormat.VDX);

        }
    }
}
