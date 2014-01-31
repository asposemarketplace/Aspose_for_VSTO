using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string mypath ="";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;
            // --- Create header for the first page. ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
              
            // Set font properties for header text.
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            // Specify header title for the first page.
            builder.Write("Header - Title Page.");

            doc.Save(mypath + "Insert Headers and Footers.doc");
        }
    }
}
