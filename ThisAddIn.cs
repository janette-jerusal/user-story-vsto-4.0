using System;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace UserStorySimilarityAddIn
{
    public partial class ThisAddIn
    {
        private MyRibbon ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Nothing required here
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Clean-up code here
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
