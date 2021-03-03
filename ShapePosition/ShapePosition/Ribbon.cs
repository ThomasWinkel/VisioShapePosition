using Microsoft.Office.Tools.Ribbon;

namespace ShapePosition
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnPositionExcel_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.EditPositionsInExcel();
        }

        private void btnDuplicateInExcel_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DuplicateInExcel();
        }

        private void btnReplace_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Replace();
        }
    }
}
