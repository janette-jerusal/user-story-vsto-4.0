using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Core;

namespace UserStorySimilarityAddInFixed
{
    public partial class ThisAddIn
    {
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e) { }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
    }
}
