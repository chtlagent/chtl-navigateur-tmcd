using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace CHTL_Navigateur_TMCD
{
    public partial class ThisAddIn
    {
        private readonly Dictionary<int, CustomTaskPane> panes = new Dictionary<int, CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WorkbookActivate += Application_WorkbookActivate;
            if (TmcdPane.IsTargetWorkbook(Application.ActiveWorkbook)) OpenTmcdPane();
        }

        private void Application_WorkbookActivate(Excel.Workbook wb)
        {
            if (TmcdPane.IsTargetWorkbook(wb)) OpenTmcdPane();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WorkbookActivate -= Application_WorkbookActivate;
            panes.Clear();
        }

        public void OpenTmcdPane()
        {
            Excel.Workbook wb = Application.ActiveWorkbook;
            if (!TmcdPane.IsTargetWorkbook(wb))
            {
                MessageBox.Show("Ouvrez un classeur contenant DOT_Base et DOT - Contrôle.",
                    "Navigateur TMCD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Excel.Window wnd = Application.ActiveWindow;
            if (wnd == null) return;
            int key = wnd.Hwnd;
            CustomTaskPane pane;
            if (!panes.TryGetValue(key, out pane))
            {
                pane = CustomTaskPanes.Add(new TmcdPane(), "CHTL | Navigateur TMCD", wnd);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = 365;
                panes[key] = pane;
            }
            ((TmcdPane)pane.Control).RefreshSheets();
            pane.Visible = true;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new TmcdRibbon();
        }

        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}