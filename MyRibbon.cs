using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace UserStorySimilarityAddInFixed
{
    public class MyRibbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public MyRibbon()
        {
            MessageBox.Show("Ribbon Constructor Called");
        }

        public string GetCustomUI(string ribbonID)
        {
            MessageBox.Show("GetCustomUI triggered!");
            var assembly = Assembly.GetExecutingAssembly();
            var resources = assembly.GetManifestResourceNames();
            MessageBox.Show("Resources:\n" + string.Join("\n", resources));

            foreach (string res in resources)
            {
                if (res.EndsWith("CustomRibbon.xml"))
                {
                    using (Stream stream = assembly.GetManifestResourceStream(res))
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        var xml = reader.ReadToEnd();
                        MessageBox.Show("Ribbon XML content:\n" + xml);
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
