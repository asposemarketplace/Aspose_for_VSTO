using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace VSTO_Diagram
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Create Visio Application Object
            Visio.Application vsdApp = Application;

            //Make Visio Application Invisible
            vsdApp.Visible = false;

            //Create a document object and load a diagram
            Visio.Document vsdDoc = vsdApp.Documents.Open("Drawing.vsd");

            //Save the VDX diagram
            vsdDoc.SaveAs("Drawing1.vdx");

            //Save as PDF file
            vsdDoc.ExportAsFixedFormat(Visio.VisFixedFormatTypes.visFixedFormatPDF,
                "Drawing1.pdf", Visio.VisDocExIntent.visDocExIntentScreen,
                Visio.VisPrintOutRange.visPrintAll, 1, vsdDoc.Pages.Count, false, true,
                true, true, true, System.Reflection.Missing.Value);

            Visio.Page vsdPage = vsdDoc.Pages[1];

            //Save as JPEG Image
            vsdPage.Export("Drawing1.jpg");

            //Quit Visio Object
            vsdApp.Quit();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
