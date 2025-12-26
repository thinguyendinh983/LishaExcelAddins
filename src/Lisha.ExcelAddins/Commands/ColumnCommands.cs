using ExcelDna.Integration;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lisha.ExcelAddins.Commands
{
    public static class ColumnCommands
    {
        [ExcelCommand(Name = "MoveColumnsLeftCommand", Description = "Moves the selected columns(s) left by one position.")]
        public static void MoveColumnsLeft()
        {
            // Get the Excel Application object using Excel-DNA
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            // Ensure a range is selected
            if (xlApp.Selection is not Excel.Range selectedRange || selectedRange.Columns.Count < 1) return;

            // Get the first column of the selection
            int firstColumnIndex = selectedRange.Column;
            int columnCount = selectedRange.Columns.Count;

            // Check if it's possible to move further left (must be at least column 2)
            if (firstColumnIndex <= 1) return;

            try
            {
                // Select and Cut the source columns
                // We select the entire columns to ensure the move works correctly
                Excel.Range sourceColumns = xlApp.Range[selectedRange.Columns[1], selectedRange.Columns[columnCount]];

                sourceColumns.Cut();

                // Define the destination range (one column to the left of the original start)
                Excel.Range destinationColumn = xlApp.Columns[firstColumnIndex - 1];

                // Insert the cut cells
                // This shifts existing data to the right
                destinationColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                // Reselect the moved rows to give feedback to the user

                // Note: The `Insert` method automatically places the cut columns at the target location.

                // Clean up COM objects (important for stability)
                Marshal.ReleaseComObject(selectedRange);
                Marshal.ReleaseComObject(sourceColumns);
                Marshal.ReleaseComObject(destinationColumn);
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi di chuyển cột: {ex.Message}");
            }
            finally
            {
                // Clear the clipboard
                xlApp.CutCopyMode = 0;
            }
        }

        [ExcelCommand(Name = "MoveColumnsRightCommand", Description = "Moves the selected columns(s) right by one position.")]
        public static void MoveColumnsRight()
        {
            throw new NotImplementedException("Chức năng di chuyển cột sang phải chưa được triển khai.");
        }
    }
}
