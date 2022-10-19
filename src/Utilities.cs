using System;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	public static class Utilities
	{
		/// <summary>
		/// Converts a column index to its column name ("A" or "AA" or whatever)
		/// </summary>
		/// <param name="idx"></param>
		/// <returns></returns>
		public static string ConvertIndexToColumnName(int idx)
		{
			if (idx <= 25)
			{
				char colName = (char)('A' + idx);
				return colName.ToString();
			}

			int n = idx / 26; // index 26 would be column "AA", index 52 would be column "BA"
			char[] chars = new char[2];
			chars[0] = (char)('A' + (byte)(n - 1));
			byte remainder = (byte)(idx % 26);
			chars[1] = (char)('A' + remainder);
			return new string(chars);
		}

		/// <summary>
		/// Creates a string representing a range of cells in a single column, e.g., A1:A14
		/// </summary>
		/// <param name="columnName">Letter indicating the column</param>
		/// <param name="startRowNum">The starting row number</param>
		/// <param name="endRowNum">The ending row number</param>
		/// <param name="options">Options to fix the column and/or row</param>
		/// <param name="sheet">Sheet name (optional)</param>
		/// <returns></returns>
		public static string CreateCellRangeString(string columnName, int startRowNum, int endRowNum, CellRangeOptions options = CellRangeOptions.None, string sheet = null)
			=> CreateCellRangeString(columnName, startRowNum, columnName, endRowNum, options, sheet);

		/// <summary>
		/// Creates a string representing a range of cells, e.g., A1:K14
		/// </summary>
		/// <param name="startColumn">Letter indicating the first column of the range</param>
		/// <param name="startRowNum">The starting row number</param>
		/// <param name="endColumn">Letter indicating the last column of the range</param>
		/// <param name="endRowNum">The ending row number</param>
		/// <param name="options">Options to fix the column and/or row</param>
		/// <param name="sheet">Sheet name (optional)</param>
		/// <returns></returns>
		public static string CreateCellRangeString(string startColumn, int startRowNum, string endColumn, int endRowNum, CellRangeOptions options = CellRangeOptions.None, string sheet = null)
		{
			bool fixedRow = (options & CellRangeOptions.FixRow) == CellRangeOptions.FixRow;
			bool fixedColumn = (options & CellRangeOptions.FixColumn) == CellRangeOptions.FixColumn;

			string rowPrefix = fixedRow ? "$" : string.Empty;
			string colPrefix = fixedColumn ? "$" : string.Empty;

			return CreateCellRangeString($"{colPrefix}{startColumn}{rowPrefix}{startRowNum}", $"{colPrefix}{endColumn}{rowPrefix}{endRowNum}", sheet);
		}

		/// <summary>
		/// Creates a string representing a range of cells, with or without another sheet name, e.g., A1:K14, 'Sheet 1'!A1:K14
		/// </summary>
		/// <param name="startCell"></param>
		/// <param name="endCell"></param>
		/// <param name="sheet">Sheet name (optional)</param>
		/// <returns></returns>
		public static string CreateCellRangeString(string startCell, string endCell, string sheet = null)
		{
			string sheetRef = CreateSheetReference(sheet);
			return $"{sheetRef}{startCell}:{endCell}";
		}

		/// <summary>
		/// Creates a string representing a cell reference, with or without another sheet name, e.g., A1, 'Sheet 1'!A1
		/// </summary>
		/// <param name="column">Letter indicating the column</param>
		/// <param name="row">Row number</param>
		/// <param name="sheet">Sheet name (optional)</param>
		/// <returns></returns>
		public static string CreateCellReference(string column, int row, string sheet = null)
		{
			string sheetRef = CreateSheetReference(sheet);
			return $"{sheetRef}{column}{row}";
		}

		private static string CreateSheetReference(string sheet) => sheet == null ? string.Empty : $"'{sheet}'!";

		/// <summary>
		/// Creates an <see cref="UpdateRequest"/> for a "header label": a row that spans multiple columns and has only one cell that has text in it
		/// </summary>
		/// <param name="sheetName">The name of the sheet</param>
		/// <param name="rowIndex">The index of the row</param>
		/// <param name="endColumnIndex">The index of the last column that the header spans across</param>
		/// <param name="label">The text of the header</param>
		/// <param name="offset">The number of cells to leave blank to the left of the label (default: 0)</param>
		/// <param name="formatter">An action to format the cell, such as by setting the background color, text color, and/or text formatting (optional)</param>
		/// <returns></returns>
		public static UpdateRequest CreateHeaderLabelRowRequest(string sheetName, int rowIndex, int endColumnIndex, string label, int offset, Action<GoogleSheetCell> formatter = null)
		{
			UpdateRequest request = new UpdateRequest(sheetName)
			{
				RowStart = rowIndex,
			};
			request.Rows.Add(CreateHeaderLabelRow(endColumnIndex, label, offset, formatter));
			return request;
		}

		/// <summary>
		/// Creates a "header label": a row that spans multiple columns and has only one cell that has text in it
		/// </summary>
		/// <param name="endColumnIndex">The index of the last column that the header spans across</param>
		/// <param name="label">The text of the header</param>
		/// <param name="offset">The number of cells to leave blank to the left of the label (default: 0)</param>
		/// <param name="formatter">An action to format the cell, such as by setting the background color, text color, and/or text formatting (optional)</param>
		/// <returns></returns>
		public static GoogleSheetRow CreateHeaderLabelRow(int endColumnIndex, string label, int offset = 0, Action<GoogleSheetCell> formatter = null)
		{
			GoogleSheetRow row = new GoogleSheetRow();
			if (offset > 0)
			{
				// put in a bunch of empty cells if needed
				for (int i = 0; i < offset; i++)
				{
					row.Add(new GoogleSheetCell(string.Empty));
				}
			}
			row.Add(new GoogleSheetCell(label));
			if (endColumnIndex > offset)
			{
				// put in empty cells after the label if needed
				for (int i = 0; i < endColumnIndex - offset; i++)
				{
					row.Add(new GoogleSheetCell(string.Empty));
				}
			}

			foreach (GoogleSheetCell cell in row)
			{
				formatter?.Invoke(cell);
			}
			return row;
		}
	}
}
