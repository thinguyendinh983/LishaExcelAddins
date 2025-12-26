using ExcelDna.Integration;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lisha.ExcelAddins.Commands
{
    public static class RowCommands
    {
        [ExcelCommand(Name = "MoveRowsUpCommand", Description = "Moves the selected row(s) up by one position.")]
        public static void MoveRowsUp()
        {
            // Access the Excel Application object via Excel-DNA
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            try
            {
                // Get the currently selected range
                Excel.Range selectedRange = xlApp.Selection;
                if (selectedRange == null) return;

                // Ensure whole rows are selected for this operation to work smoothly
                Excel.Range entireRows = selectedRange.EntireRow;
                int firstRowIndex = entireRows.Row;
                int rowCount = entireRows.Rows.Count;

                // Cannot move up from the first row
                if (firstRowIndex <= 1) return;

                // Cut the selected row(s)
                entireRows.Cut();

                // Select the row above the current top row as the destination
                // This destination is where the cut cells will be inserted
                Excel.Range destinationRange = xlApp.ActiveSheet.Rows[firstRowIndex - 1];

                // Insert the cut cells, shifting existing cells down
                destinationRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                // Reselect the moved rows to give feedback to the user
                // Note: The row index shifts after insertion
                ((Excel.Range)xlApp.Range[xlApp.Cells[firstRowIndex - 1, 1], xlApp.Cells[firstRowIndex + rowCount - 2, 1]]).EntireRow.Select();

                // Clean up COM objects (important for stability)
                Marshal.ReleaseComObject(selectedRange);
                Marshal.ReleaseComObject(entireRows);
                Marshal.ReleaseComObject(destinationRange);
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi di chuyển hàng: {ex.Message}", "Lỗi");
            }
            finally
            {
                // Clear the clipboard
                xlApp.CutCopyMode = 0;
            }
        }

        [ExcelCommand(Name = "MoveRowsDownCommand", Description = "Moves the selected row(s) down by one position.")]
        public static void MoveRowsDown()
        {
            throw new NotImplementedException("Chức năng di chuyển hàng xuống chưa được triển khai.");
        }
    }
}
