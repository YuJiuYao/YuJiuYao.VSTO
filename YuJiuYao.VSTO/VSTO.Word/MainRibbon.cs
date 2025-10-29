using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace VSTO.Word
{
    public partial class MainRibbon
    {
        private Microsoft.Office.Interop.Word.Application _application;

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _application = Globals.ThisAddIn.Application;
        }

        private void buttonShowMessageBox_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(_application.ActiveDocument.ToString());
        }
    }
}
