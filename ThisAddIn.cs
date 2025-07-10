using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public partial class ThisAddIn
    {
        private MyRibbon ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Optional: initialization logic
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Optional: cleanup logic
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new MyRibbon();
            return ribbon;
        }
    }
}

