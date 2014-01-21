using Aspose.Diagram;

namespace Aspose_Diagram
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load diagram
            Diagram vsdDiagram = new Diagram("Drawing.vsd");

            //Save the diagram as VDX
            vsdDiagram.Save("Drawing1.vdx", SaveFileFormat.VDX);

            //Save as PDF
            vsdDiagram.Save("Drawing1.pdf", SaveFileFormat.PDF);

            //Save as JPEG
            vsdDiagram.Save("Drawing1.jpg", SaveFileFormat.JPEG);
        }
    }
}
