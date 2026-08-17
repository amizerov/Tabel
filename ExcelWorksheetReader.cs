namespace Tabel;

using System.Runtime.InteropServices;

internal sealed record ExcelWorksheetData(object[,] Values, int Rows, int Columns);

internal static class ExcelWorksheetReader
{
    private const int MaxRowsPerRead = 5000;
    private const int MaxCellsPerRead = 100000;
    private const int XlFormulas = -4123;
    private const int XlPart = 2;
    private const int XlByRows = 1;
    private const int XlByColumns = 2;
    private const int XlPrevious = 2;

    public static ExcelWorksheetData Read(dynamic sheet, int? maxRows = null, int? maxColumns = null)
    {
        dynamic? cells = null;
        dynamic? lastRowCell = null;
        dynamic? lastColumnCell = null;
        dynamic? usedRange = null;
        dynamic? usedRangeRows = null;
        dynamic? usedRangeColumns = null;
        dynamic? firstCell = null;
        dynamic? endCell = null;
        dynamic? range = null;

        try
        {
            cells = sheet.Cells;
            lastRowCell = cells.Find(
                "*", Type.Missing, XlFormulas, XlPart, XlByRows, XlPrevious,
                false, Type.Missing, false);
            lastColumnCell = cells.Find(
                "*", Type.Missing, XlFormulas, XlPart, XlByColumns, XlPrevious,
                false, Type.Missing, false);

            int rows = lastRowCell == null ? 1 : Convert.ToInt32(lastRowCell.Row);
            int columns = lastColumnCell == null ? 1 : Convert.ToInt32(lastColumnCell.Column);

            // Find can occasionally return a stale boundary for filtered/imported sheets.
            // UsedRange is kept as a fallback so records below that boundary are not lost.
            usedRange = sheet.UsedRange;
            int usedFirstRow = Convert.ToInt32(usedRange.Row);
            int usedFirstColumn = Convert.ToInt32(usedRange.Column);
            usedRangeRows = usedRange.Rows;
            usedRangeColumns = usedRange.Columns;
            int usedRowCount = Convert.ToInt32(usedRangeRows.Count);
            int usedColumnCount = Convert.ToInt32(usedRangeColumns.Count);
            rows = Math.Max(rows, usedFirstRow + usedRowCount - 1);
            columns = Math.Max(columns, usedFirstColumn + usedColumnCount - 1);

            if (maxRows.HasValue)
                rows = Math.Min(rows, maxRows.Value);
            if (maxColumns.HasValue)
                columns = Math.Min(columns, maxColumns.Value);

            var values = new object[rows + 1, columns + 1];
            int rowsPerRead = Math.Max(1, Math.Min(MaxRowsPerRead, MaxCellsPerRead / columns));

            for (int firstRow = 1; firstRow <= rows; firstRow += rowsPerRead)
            {
                int lastRow = Math.Min(rows, firstRow + rowsPerRead - 1);
                firstCell = cells[firstRow, 1];
                endCell = cells[lastRow, columns];
                range = sheet.Range[firstCell, endCell];

                CopyTo(values, range.Value2, firstRow, lastRow, columns);

                ReleaseComObject(range);
                range = null;
                ReleaseComObject(endCell);
                endCell = null;
                ReleaseComObject(firstCell);
                firstCell = null;
            }

            return new ExcelWorksheetData(values, rows, columns);
        }
        finally
        {
            ReleaseComObject(range);
            ReleaseComObject(endCell);
            ReleaseComObject(firstCell);
            ReleaseComObject(lastColumnCell);
            ReleaseComObject(lastRowCell);
            ReleaseComObject(usedRangeColumns);
            ReleaseComObject(usedRangeRows);
            ReleaseComObject(usedRange);
            ReleaseComObject(cells);
        }
    }

    private static void CopyTo(object[,] destination, object value, int firstRow, int lastRow, int columns)
    {
        if (value is not object[,] source)
        {
            destination[firstRow, 1] = value;
            return;
        }

        int sourceFirstRow = source.GetLowerBound(0);
        int sourceFirstColumn = source.GetLowerBound(1);
        int rowCount = lastRow - firstRow + 1;

        for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
        {
            for (int columnOffset = 0; columnOffset < columns; columnOffset++)
            {
                destination[firstRow + rowOffset, columnOffset + 1] =
                    source[sourceFirstRow + rowOffset, sourceFirstColumn + columnOffset];
            }
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value == null)
            return;

        try { Marshal.FinalReleaseComObject(value); } catch { }
    }
}
