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
            Visio.Application vsdApp = null;
            Visio.Document vsdDoc = null;

            //Create Visio Application Object
            vsdApp = Application;

            //Make Visio Application Invisible
            vsdApp.Visible = false;

            //Create a document object and load a diagram
            vsdDoc = vsdApp.Documents.Open("Drawing.vsd");

            //Create page object to get required page
            Visio.Page page = vsdApp.ActivePage;

            //Create shape object to get required shape
            Visio.Shape shape = page.Shapes["Process1"];

            //Set shape text and text style
            shape.Text = "Hello World";
            shape.TextStyle = "CustomStyle1";

            //Set shape's position
            shape.get_Cells("PinX").ResultIU = 5;
            shape.get_Cells("PinY").ResultIU = 5;

            //Set shape's height and width
            shape.get_Cells("Height").ResultIU = 2;
            shape.get_Cells("Width").ResultIU = 3;

            //Save file as VDX
            vsdDoc.SaveAs("Drawing1.vdx");
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
