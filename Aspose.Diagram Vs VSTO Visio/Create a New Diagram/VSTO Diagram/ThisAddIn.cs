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
            Visio.Application vdxApp = null;
            Visio.Document vdxDoc = null;

            //Create Visio Application Object
            vdxApp = Application;

            //Make Visio Application Invisible
            vdxApp.Visible = false;

            //Create a new diagram
            vdxDoc = vdxApp.Documents.Add("Drawing.vsd");

            //Load Visio Stencil
            Visio.Documents visioDocs = vdxApp.Documents;
            
            Visio.Document visioStencil = visioDocs.OpenEx("sample.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenHidden);

            //Set active page
            Visio.Page visioPage = vdxApp.ActivePage;

            //Add a new rectangle shape
            Visio.Master visioRectMaster = visioStencil.Masters.get_ItemU(@"Rectangle");
            Visio.Shape visioRectShape = visioPage.Drop(visioRectMaster, 4.25, 5.5);
            visioRectShape.Text = @"Rectangle text.";

            //Add a new star shape
            Visio.Master visioStarMaster = visioStencil.Masters.get_ItemU(@"Star 7");
            Visio.Shape visioStarShape = visioPage.Drop(visioStarMaster, 2.0, 5.5);
            visioStarShape.Text = @"Star text.";

            //Add a new hexagon shape
            Visio.Master visioHexagonMaster = visioStencil.Masters.get_ItemU(@"Hexagon");
            Visio.Shape visioHexagonShape = visioPage.Drop(visioHexagonMaster, 7.0, 5.5);
            visioHexagonShape.Text = @"Hexagon text.";


            //Save diagram as VDX
            vdxDoc.SaveAs("Drawing1.vdx");
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
