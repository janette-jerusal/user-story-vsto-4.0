using System.Reflection;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace UserStorySimilarityAddIn
{
    public class MyRibbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public MyRibbon()
        {
            // Confirm constructor runs at startup
            MessageBox.Show("Ribbon Constructor Called");
        }

        public string GetCustomUI(string ribbonID)
        {
            MessageBox.Show("GetCustomUI triggered!");

            var assembly = Assembly.GetExecutingAssembly();
            var resources = assembly.GetManifestResourceNames();
            MessageBox.Show("Embedded Resources:\n" + string.Join("\n", resources));

            foreach (var res in resources)
            {
                if (res.EndsWith("MvRibbon.xml", System.StringComparison.OrdinalIgnoreCase))
                {
                    using (Stream stream = assembly.GetManifestResourceStream(res))
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string xml = reader.ReadToEnd();
                        MessageBox.Show("Ribbon XML loaded:\n" + xml.Substring(0, Math.Min(xml.Length, 200)));
                        return xml;
                    }
                }
            }

            MessageBox.Show("MvRibbon.xml not found!");
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
