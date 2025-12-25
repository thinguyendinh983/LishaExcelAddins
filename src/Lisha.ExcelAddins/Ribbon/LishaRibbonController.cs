using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Lisha.ExcelAddins.Ribbon
{
    [ComVisible(true)]
    public class LishaRibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return LishaResource.LishaRibbon;
        }

        public override object? LoadImage(string imageId)
        {
            // This will return the image resource with the name specified in the image='xxxx' tag
            return LishaResource.ResourceManager.GetObject(imageId);
        }
        public void OnbtnNumToStringPressed(IRibbonControl control)
        {
            // Get the Excel Application object via Excel-DNA
            Application xlApp = (Application)ExcelDnaUtil.Application;

            // Access the active cell range
            Range activeCell = xlApp.ActiveCell;

            // Set the value of the active cell
            activeCell.Value = Functions.LishaFunctions.LishaSoSangChu(activeCell.Value, true);
        }

        public void OnbtnAboutPressed(IRibbonControl control)
        {
            MessageBox.Show("Lisha Excel Add-in\nVersion 1.0.0\n\nDeveloped by Nguyen Dinh Thi\nEmail: thinguyendinh983@gmail.com",
                "About Lisha Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
