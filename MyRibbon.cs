using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace UserStorySimilarityAddIn
{
    [ComVisible(true)]
    public class MyRibbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var res = typeof(MyRibbon).Assembly.GetManifestResourceNames();
            foreach (var r in res)
            {
                if (r.EndsWith("CustomRibbon.xml"))
                {
                    using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(r))
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string xml = reader.ReadToEnd();
                        MessageBox.Show("Ribbon XML loaded:\n" + xml);
                        return xml;
                    }
                }
            }

            MessageBox.Show("CustomRibbon.xml not found in resources.");
            return null;
        }

        public void OnLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnCompareClick(IRibbonControl control)
        {
            MessageBox.Show("Compare button clicked!");
        }
    }
}
