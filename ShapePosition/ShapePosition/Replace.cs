using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Visio = Microsoft.Office.Interop.Visio;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ShapePosition
{
    class Replace
    {
        Visio.Application vApp;
        Visio.Window vWin;
        Visio.Document vDoc;

        Excel.ListObject listObject;

        public Replace()
        {
            vApp = Globals.ThisAddIn.Application;
            vWin = vApp.ActiveWindow;
            vDoc = vApp.ActiveDocument;

            if (vWin.Type != (short)Visio.VisWinTypes.visDrawing) return;

            Excel.Application xlsApp = new Excel.Application();
            Excel.Workbook workbook = xlsApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets[1];

            worksheet.Name = "Replace";

            worksheet.Cells[1, 1] = "No";
            worksheet.Cells[1, 2] = "Page";
            worksheet.Cells[1, 3] = "Id";
            worksheet.Cells[1, 4] = "Text";
            worksheet.Cells[1, 5] = "Replace";

            listObject = worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]], null, Excel.XlYesNoGuess.xlYes);

            if (vWin.Selection.Count > 0)
            {
                for (int i = 1; i <= vWin.Selection.Count; i++)
                {
                    Visio.Shape shape = vWin.Selection[i];
                    if (shape.Text == "") continue;
                    worksheet.Cells[i + 1, 1] = i;
                    worksheet.Cells[i + 1, 2] = shape.ContainingPage.Name;
                    worksheet.Cells[i + 1, 3] = shape.ID;
                    worksheet.Cells[i + 1, 4] = shape.Text;
                    worksheet.Cells[i + 1, 5] = shape.Text;
                }
            }
            else
            {
                foreach(Visio.Page page in vDoc.Pages)
                {
                    for (int i = 1; i <= page.Shapes.Count; i++)
                    {
                        Visio.Shape shape = page.Shapes[i];
                        if (shape.Text == "") continue;
                        worksheet.Cells[i + 1, 1] = i;
                        worksheet.Cells[i + 1, 2] = shape.ContainingPage.Name;
                        worksheet.Cells[i + 1, 3] = shape.ID;
                        worksheet.Cells[i + 1, 4] = shape.Text;
                        worksheet.Cells[i + 1, 5] = shape.Text;
                    }
                }
            }

            worksheet.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(Worksheet_SelectionChange);
            xlsApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(App_WorkbookBeforeClose);


            worksheet.Columns.AutoFit();
            xlsApp.Visible = true;
        }

        public void App_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            DialogResult dialogResult = MessageBox.Show("Replace shape text?", "Replace", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                foreach (Excel.ListRow listRow in listObject.ListRows)
                {
                    string textOld = listRow.Range[1, 4].Value2.ToString();
                    string textNew = listRow.Range[1, 5].Value2.ToString();

                    if (textOld == textNew) continue;

                    string page = listRow.Range[1, 2].Value2.ToString();
                    int id = int.Parse(listRow.Range[1, 3].Value2.ToString());
                    Visio.Shape shape = vApp.ActiveDocument.Pages[page].Shapes.ItemFromID[id];

                    shape.Text = textNew;
                }
            }

            workbook.Saved = true;
            cancel = false;
        }

        public void Worksheet_SelectionChange(Excel.Range target)
        {
            try
            {
                Excel.Worksheet worksheet = target.Worksheet;

                string page = (worksheet.Cells[target.Row, 2] as Excel.Range).Value2.ToString();
                int id = int.Parse((worksheet.Cells[target.Row, 3] as Excel.Range).Value2.ToString());

                Visio.Shape shape = vApp.ActiveDocument.Pages[page].Shapes.ItemFromID[id];
                vApp.ActiveWindow.CenterViewOnShape(shape, Visio.VisCenterViewFlags.visCenterViewSelectShape);
            }
            catch
            {
                vApp.ActiveWindow.DeselectAll();
            }
        }
    }
}