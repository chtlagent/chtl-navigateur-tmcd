using Office = Microsoft.Office.Core;
namespace CHTL_Navigateur_TMCD
{
    public sealed class TmcdRibbon : Office.IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui""><ribbon><tabs><tab idMso=""TabAddIns""><group id=""chtlNavGroup"" label=""CHTL | Pilotage""><button id=""openCHTLTMCD"" label=""Navigateur TMCD"" size=""large"" imageMso=""FileOpen"" onAction=""OpenTMCD"" /></group></tab></tabs></ribbon></customUI>";
        }
        public void OpenTMCD(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.OpenTmcdPane();
        }
    }
}