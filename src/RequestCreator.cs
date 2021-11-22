using System;
using System.Collections.Generic;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	public static class RequestCreator
	{
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
