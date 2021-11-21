using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	public static class RequestCreator
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
		/// <returns></returns>
		public static string CreateCellRangeString(string columnName, int startRowNum, int endRowNum, CellRangeOptions options = CellRangeOptions.None)
		{
			return CreateCellRangeString(columnName, startRowNum, columnName, endRowNum, options);
		}

		/// <summary>
		/// Creates a string representing a range of cells, e.g., A1:K14
		/// </summary>
		/// <param name="startColumn">Letter indicating the first column of the range</param>
		/// <param name="startRowNum">The starting row number</param>
		/// <param name="endColumn">Letter indicating the last column of the range</param>
		/// <param name="endRowNum">The ending row number</param>
		/// <param name="options">Options to fix the column and/or row</param>
		/// <returns></returns>
		public static string CreateCellRangeString(string startColumn, int startRowNum, string endColumn, int endRowNum, CellRangeOptions options = CellRangeOptions.None)
		{
			bool fixedRow = (options & CellRangeOptions.FixRow) == CellRangeOptions.FixRow;
			bool fixedColumn = (options & CellRangeOptions.FixColumn) == CellRangeOptions.FixColumn;

			string rowPrefix = fixedRow ? "$" : string.Empty;
			string colPrefix = fixedColumn ? "$" : string.Empty;

			return $"{colPrefix}{startColumn}{rowPrefix}{startRowNum}:{colPrefix}{endColumn}{rowPrefix}{endRowNum}";
		}

		/// <summary>
		/// Creates a <see cref="SetDataValidationRequest"/> for drop-down menus to choose from a range of values
		/// </summary>
		/// <param name="sourceSheetName">Name of the sheet where the values are</param>
		/// <param name="firstSourceCell">The first cell of the value range (e.g., A1)</param>
		/// <param name="lastSourceCell">The last cell of the value range (e.g., A10)</param>
		/// <param name="sheetId">ID of the sheet where you want to put the validation</param>
		/// <param name="startRowIndex">Index of the row where validation starts</param>
		/// <param name="startColumnIndex">Index of the column where validation starts</param>
		/// <param name="offset">Number of rows or columns that are affected by this request</param>
		/// <param name="direction">Indicates direction that the validation is applied</param>
		/// <returns></returns>
		public static Request CreateDataValidationRequest(string sourceSheetName, string firstSourceCell, string lastSourceCell, int? sheetId, int startRowIndex, int startColumnIndex, int offset, RepeatDirection direction = RepeatDirection.Vertical)
		{
			if (offset < 1)
				throw new ArgumentException("Must be >= 1", nameof(offset));
			return new Request
			{
				SetDataValidation = new SetDataValidationRequest
				{
					Range = new GridRange
					{
						SheetId = sheetId,
						StartRowIndex = startRowIndex,
						StartColumnIndex = startColumnIndex,
						EndRowIndex = startRowIndex + (direction == RepeatDirection.Vertical ? offset : 1),
						EndColumnIndex = startColumnIndex + (direction == RepeatDirection.Horizontal ? offset : 1),
					},
					Rule = new DataValidationRule
					{
						Condition = new BooleanCondition
						{
							Type = "ONE_OF_RANGE",
							Values = new List<ConditionValue>
							{
								new ConditionValue
								{
									UserEnteredValue = $"={sourceSheetName}!{firstSourceCell}:{lastSourceCell}",
								},
							},
						},
						ShowCustomUi = true,
						Strict = true,
					},
				},
			};
		}

		/// <summary>
		/// Creates a <see cref="RepeatCellRequest"/> for a formula to be inserted across a range
		/// </summary>
		/// <param name="sheetId">ID of the sheet where you want to put the formula</param>
		/// <param name="startRowIndex">Index of the row where the formula starts</param>
		/// <param name="startColumnIndex">Index of the column where formula is to be applied</param>
		/// <param name="rowCount">Number of rows affected</param>
		/// <param name="formula">Tne formula to repeat</param>
		/// <returns></returns>
		public static Request CreateRepeatedSheetFormulaRequest(int? sheetId, int startRowIndex, int startColumnIndex, int rowCount, string formula, RepeatDirection direction = RepeatDirection.Vertical)
		{
			return new Request
			{
				RepeatCell = new RepeatCellRequest
				{
					Range = new GridRange
					{
						SheetId = sheetId,
						StartRowIndex = startRowIndex,
						EndRowIndex = startRowIndex + rowCount,
						StartColumnIndex = startColumnIndex,
						EndColumnIndex = startColumnIndex + 1,
					},
					Cell = new CellData
					{
						UserEnteredValue = new ExtendedValue
						{
							FormulaValue = formula,
						}
					},
					Fields = "userEnteredValue",
				},
			};
		}

		/// <summary>
		/// Creates a <see cref="RepeatCellRequest"/> to format a range of cells (assumes you want to start from the first column)
		/// </summary>
		/// <param name="sheetId"></param>
		/// <param name="startRowIndex"></param>
		/// <param name="endColumnIndex"></param>
		/// <param name="cellDataAction"></param>
		/// <returns></returns>
		public static Request CreateRowFormattingRequest(int? sheetId, int startRowIndex, int endColumnIndex, Func<CellData> cellDataAction)
		{
			return CreateRowFormattingRequest(sheetId, startRowIndex, 0, endColumnIndex, cellDataAction);
		}

		/// <summary>
		/// Creates a <see cref="RepeatCellRequest"/> to format a range of cells
		/// </summary>
		/// <param name="sheetId"></param>
		/// <param name="startRowIndex"></param>
		/// <param name="startColumnIndex"></param>
		/// <param name="endColumnIndex"></param>
		/// <param name="cellDataAction"></param>
		/// <returns></returns>
		public static Request CreateRowFormattingRequest(int? sheetId, int startRowIndex, int startColumnIndex, int endColumnIndex, Func<CellData> cellDataAction)
		{
			return new Request
			{
				RepeatCell = new RepeatCellRequest
				{
					Range = new GridRange
					{
						SheetId = sheetId,
						StartRowIndex = startRowIndex,
						StartColumnIndex = startColumnIndex,
						EndRowIndex = startRowIndex + 1,
						EndColumnIndex = endColumnIndex + 1,
					},
					Cell = cellDataAction(),
					Fields = "userEnteredFormat(backgroundColor,textFormat)",
				},
			};
		}
	}
}
