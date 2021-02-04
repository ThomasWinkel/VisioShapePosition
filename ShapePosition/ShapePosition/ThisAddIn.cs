using System;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Excel = Microsoft.Office.Interop.Excel;

namespace ShapePosition
{
    public partial class ThisAddIn
    {
        Visio.Application vApp;

        public void EditPositionsInExcel()
        {
            Visio.Window win = vApp.ActiveWindow;
            if (win.Type != (short)Visio.VisWinTypes.visDrawing) return;
            if (win.Selection.Count < 1) return;

            Excel.Application xlsApp = new Excel.Application();
            Excel.Workbook workbook = xlsApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            worksheet.Name = "Positions";
            worksheet.Cells[1, 1] = "Page";
            worksheet.Cells[1, 2] = "Id";
            worksheet.Cells[1, 3] = "X";
            worksheet.Cells[1, 4] = "Y";
            Excel.ListObject listObject = worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 4]], null, Excel.XlYesNoGuess.xlYes);

            for(int i = 1; i <= vApp.ActiveWindow.Selection.Count; i++)
            {
                Visio.Shape shape = vApp.ActiveWindow.Selection[i];
                worksheet.Cells[i + 1, 1] = shape.ContainingPage.Name;
                worksheet.Cells[i + 1, 2] = shape.ID;
                worksheet.Cells[i + 1, 3] = shape.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters];
                worksheet.Cells[i + 1, 4] = shape.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters];
            }

            worksheet.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(Worksheet_SelectionChange);
            worksheet.Change += new Excel.DocEvents_ChangeEventHandler(Worksheet_Change);

            worksheet.Columns.AutoFit();
            xlsApp.Visible = true;
        }

        public void DuplicateInExcel()
        {
            Visio.Window win = vApp.ActiveWindow;
            if (win.Type != (short)Visio.VisWinTypes.visDrawing) return;
            if (win.Selection.Count < 1) return;

            Excel.Application xlsApp = new Excel.Application();
            Excel.Workbook workbook = xlsApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            worksheet.Name = "Positions";
            worksheet.Cells[1, 1] = "X";
            worksheet.Cells[1, 2] = "Y";
            Excel.ListObject listObject = worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 2]], null, Excel.XlYesNoGuess.xlYes);

            xlsApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(App_WorkbookBeforeClose);

            worksheet.Columns.AutoFit();
            xlsApp.Visible = true;
        }

        public void App_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            Excel.ListObject listObject = workbook.Worksheets[1].ListObjects[1];
            Visio.Shape shapePrimary = vApp.ActiveWindow.Selection.PrimaryItem;

            foreach(Excel.ListRow listRow in listObject.ListRows)
            {
                if (double.TryParse(listRow.Range[1, 1].Value2.ToString(), out double posX))
                {
                    if (double.TryParse(listRow.Range[1, 2].Value2.ToString(), out double posY))
                    {
                        Visio.Shape shape = shapePrimary.Duplicate();
                        shape.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = posX;
                        shape.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = posY;
                    }
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

                string page = (worksheet.Cells[target.Row, 1] as Excel.Range).Value2.ToString();
                int id = int.Parse((worksheet.Cells[target.Row, 2] as Excel.Range).Value2.ToString());

                Visio.Shape shape = vApp.ActiveDocument.Pages[page].Shapes.ItemFromID[id];
                vApp.ActiveWindow.CenterViewOnShape(shape, Visio.VisCenterViewFlags.visCenterViewSelectShape);
            }
            catch
            {
                vApp.ActiveWindow.DeselectAll();
            }
        }

        public void Worksheet_Change(Excel.Range target)
        {
            try
            {
                Excel.Worksheet worksheet = target.Worksheet;

                for(int i = 1; i <= target.Count; i++)
                {
                    Excel.Range cell = target[i];

                    string page = (worksheet.Cells[cell.Row, 1] as Excel.Range).Value2.ToString();
                    int id = int.Parse((worksheet.Cells[cell.Row, 2] as Excel.Range).Value2.ToString());

                    Visio.Shape shape = vApp.ActiveDocument.Pages[page].Shapes.ItemFromID[id];

                    if (double.TryParse(cell.Value2.ToString(), out double pos))
                    {
                        if (cell.Column == 3)
                        {
                            shape.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = pos;
                        }
                        else if (cell.Column == 4)
                        {
                            shape.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = pos;
                        }
                    }
                }
            }
            catch { }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            vApp = Globals.ThisAddIn.Application;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }


        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

    }
}
