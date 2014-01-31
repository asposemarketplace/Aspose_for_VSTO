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
            vsdDoc = vsdApp.Documents.Open("Add Shapes.vdx");

            Visio.Documents visioDocs = this.Application.Documents;
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);

            Visio.Page visioPage = this.Application.ActivePage;

            Visio.Master visioRectMaster = visioStencil.Masters.get_ItemU(@"Rectangle");
            Visio.Shape visioRectShape = visioPage.Drop(visioRectMaster, 4.25, 5.5);
            visioRectShape.Text = @"Rectangle text.";

            Visio.Master visioStarMaster = visioStencil.Masters.get_ItemU(@"Star 7");
            Visio.Shape visioStarShape = visioPage.Drop(visioStarMaster, 2.0, 5.5);
            visioStarShape.Text = @"Star text.";

            Visio.Master visioHexagonMaster = visioStencil.Masters.get_ItemU(@"Hexagon");
            Visio.Shape visioHexagonShape = visioPage.Drop(visioHexagonMaster, 7.0, 5.5);
            visioHexagonShape.Text = @"Hexagon text.";
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
