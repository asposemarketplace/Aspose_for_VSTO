using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Diagram;

namespace Aspose_Diagram
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load masters from any existing diagram, stencil or template
            // and add in the new diagram
            string visioStencil = "Add Shapes.vdx";

            //Names of the masters present in the stencil
            string rectangleMaster = @"Rectangle";

            int pageNumber = 0;
            double width = 2, height = 2, pinX = 4.25, pinY = 9.5;

            // Create a new diagram
            Diagram diagram = new Diagram(visioStencil);

            //Add a new rectangle shape
            long rectangleId = diagram.AddShape(
                pinX, pinY, width, height, rectangleMaster, pageNumber);
            
            //Save the diagram
            diagram.Save("Add Shapes.vdx", SaveFileFormat.VDX);

        }
    }
}
