using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CHTL_Navigateur_TMCD
{
    public sealed class TmcdPane : UserControl
    {
        private readonly TreeView tree = new TreeView();
        private readonly TextBox search = new TextBox();
        private readonly CheckBox showHidden = new CheckBox();
        private readonly Label status = new Label();
        private string previousSheet;
        private string previousBook;

        private static readonly KeyValuePair<string,string[]>[] Groups = {
            new KeyValuePair<string,string[]>("TMCD annuels", new [] {"TMCD 2026","TMCD 2025","TMCD 2024","TMCD 2023"}),
            new KeyValuePair<string,string[]>("Dotations", new [] {"b. Dotations >>>","DOT_Base","DOT - Règles","DOT - Contrôle","FIR","MIG","AC"}),
            new KeyValuePair<string,string[]>("Pilotage", new [] {"P&L","P&L CHTL","TMCD annuels","BCA","GHS & MO"}),
            new KeyValuePair<string,string[]>("GHS, MO & Journées", new [] {"R - GHS AM","R - MO","R - GHS T2A","R - Journées","a. GHS & MO >>>"}),
            new KeyValuePair<string,string[]>("Chambres individuelles", new [] {"c. Ch. ind >>>","ICR Bloc 24-25","ICR Imagerie 24-25","ICR Anesth 24-25"}),
            new KeyValuePair<string,string[]>("RH & FICOM", new [] {"Matrice HR 2026","Matrice HR 2025","Matrice HR 2024","Matrice HR 2023","FICOM"})
        };

        public TmcdPane()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10), ColumnCount = 1, RowCount = 5 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 115));
            Controls.Add(layout);
            var heading = new Label { Dock = DockStyle.Fill, BackColor = Color.FromArgb(21,35,68),
                ForeColor = Color.White, Text = "CHTL\nNAVIGATEUR TMCD", Font = new Font("Segoe UI", 15, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter };
            layout.Controls.Add(heading, 0, 0);
            var searchBlock = new Panel { Dock = DockStyle.Fill };
            search.Dock = DockStyle.Bottom;
            search.Font = new Font("Segoe UI", 11);
            search.PlaceholderFallback();
            search.TextChanged += (s,e) => RefreshSheets();
            searchBlock.Controls.Add(search);
            searchBlock.Controls.Add(new Label { Dock = DockStyle.Top, Height = 18, Text = "Rechercher une feuille",
                ForeColor = Color.FromArgb(21,111,136), Font = new Font("Segoe UI", 8, FontStyle.Bold) });
            layout.Controls.Add(searchBlock, 0, 1);
            tree.Dock = DockStyle.Fill;
            tree.Font = new Font("Segoe UI", 10);
            tree.ItemHeight = 24;
            tree.HideSelection = false;
            tree.NodeMouseDoubleClick += (s,e) => OpenNode(e.Node);
            tree.KeyDown += (s,e) => { if (e.KeyCode == Keys.Enter) { OpenNode(tree.SelectedNode); e.Handled = true; } };
            layout.Controls.Add(tree, 0, 2);
            showHidden.Text = "Afficher les feuilles masquées (sans les ouvrir)";
            showHidden.Dock = DockStyle.Fill;
            showHidden.CheckedChanged += (s,e) => RefreshSheets();
            layout.Controls.Add(showHidden, 0, 3);
            var bottom = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 2 };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            bottom.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            bottom.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            bottom.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            AddButton(bottom, "Accueil", (s,e)=>GoToSheet("P&L"), 0, 0);
            AddButton(bottom, "Précédent", (s,e)=>GoBack(), 1, 0);
            AddButton(bottom, "Contrôle DOT", (s,e)=>GoToSheet("DOT - Contrôle"), 0, 1);
            AddButton(bottom, "Effacer filtres", (s,e)=>RunMacro("Reinitialiser_Filtres_DOT"), 1, 1);
            status.Dock = DockStyle.Fill;
            status.Font = new Font("Segoe UI", 8);
            status.Text = "Local, aucune donnée envoyée.";
            bottom.Controls.Add(status,0,2);
            bottom.SetColumnSpan(status,2);
            layout.Controls.Add(bottom, 0, 4);
            RefreshSheets();
        }
        private static void AddButton(TableLayoutPanel target, string text, EventHandler action, int col, int row)
        {
            var b = new Button { Dock = DockStyle.Fill, Text = text, Font = new Font("Segoe UI", 9, FontStyle.Bold),
                BackColor = Color.FromArgb(236,243,249), FlatStyle = FlatStyle.Flat };
            b.Click += action;
            target.Controls.Add(b, col, row);
        }

        public static bool IsTargetWorkbook(Excel.Workbook wb)
        {
            if (wb == null) return false;
            try {
                return wb.Worksheets["DOT_Base"] != null && wb.Worksheets["DOT - Contrôle"] != null;
            } catch { return false; }
        }

        private Excel.Workbook Book()
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (!IsTargetWorkbook(wb)) { status.Text = "Ouvrez le classeur TMCD."; return null; }
            return wb;
        }
        public void RefreshSheets()
        {
            if (tree == null || status == null || search == null) return;
            Excel.Workbook wb = Book();
            if (wb == null) return;
            tree.BeginUpdate();
            try {
                tree.Nodes.Clear();
                var vis = new Dictionary<string,Excel.XlSheetVisibility>(StringComparer.OrdinalIgnoreCase);
                foreach (Excel.Worksheet ws in wb.Worksheets) vis[ws.Name] = ws.Visible;
                var added = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var group in Groups) {
                    var parent = new TreeNode(group.Key);
                    foreach (string name in group.Value) AddNode(parent,name,vis,added);
                    if (parent.Nodes.Count > 0) { parent.Expand(); tree.Nodes.Add(parent); }
                }
                var others = new TreeNode("Autres feuilles");
                foreach (string name in vis.Keys.OrderBy(x=>x)) {
                    if (!added.Contains(name)) AddNode(others,name,vis,added);
                }
                if (others.Nodes.Count > 0) { others.Expand(); tree.Nodes.Add(others); }
                status.Text = vis.Count + " feuilles";
            } catch (Exception ex) { status.Text = ex.Message; }
            finally { tree.EndUpdate(); }
        }
        private void AddNode(TreeNode parent,string name,Dictionary<string,Excel.XlSheetVisibility> vis,HashSet<string> added)
        {
            Excel.XlSheetVisibility v;
            if (!vis.TryGetValue(name,out v)) return;
            if (v != Excel.XlSheetVisibility.xlSheetVisible && !showHidden.Checked) return;
            if (search.TextLength > 0 && name.IndexOf(search.Text,StringComparison.CurrentCultureIgnoreCase)<0) return;
            var node = new TreeNode(name+(v==Excel.XlSheetVisibility.xlSheetVisible?"":" [masquée]")) { Tag=name };
            if (v != Excel.XlSheetVisibility.xlSheetVisible) node.ForeColor=Color.Gray;
            parent.Nodes.Add(node); added.Add(name);
        }
        private void OpenNode(TreeNode n) { if(n != null && n.Tag is string) GoToSheet((string)n.Tag); }
        private void GoToSheet(string name)
        {
            Excel.Workbook wb = Book(); if(wb==null)return;
            try {
                Excel.Worksheet dest = wb.Worksheets[name] as Excel.Worksheet;
                if (dest==null) return;
                if(dest.Visible != Excel.XlSheetVisibility.xlSheetVisible) {
                    MessageBox.Show("Feuille masquée : aucune modification de visibilité automatique.");
                    return;
                }
                Excel.Worksheet before = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                if(before != null && before.Name != name) { previousBook=wb.FullName; previousSheet=before.Name; }
                dest.Activate(); status.Text = name;
            } catch(Exception ex) { MessageBox.Show("Navigation impossible : "+ex.Message); }
        }
        private void GoBack()
        {
            Excel.Workbook wb=Book(); if(wb==null)return;
            if(previousSheet==null || !string.Equals(previousBook,wb.FullName,StringComparison.OrdinalIgnoreCase))return;
            var dest=previousSheet; previousSheet=null; GoToSheet(dest);
        }
        private void RunMacro(string name)
        {
            Excel.Workbook wb=Book(); if(wb==null)return;
            try {
                Globals.ThisAddIn.Application.Run("'"+wb.Name.Replace("'","''")+"'!"+name);
                status.Text="Filtres réinitialisés";
            } catch(Exception ex) { MessageBox.Show("Macro non exécutée : "+ex.Message); }
        }
    }
    internal static class CompatibilityExtensions
    {
        // .NET Framework WinForms does not support TextBox.PlaceholderText.
        public static void PlaceholderFallback(this TextBox tb) { }
    }
}