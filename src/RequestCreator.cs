using System;
using System.Collections.Generic;
using System.Linq;
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
									UserEnteredValue = $"={Utilities.CreateCellRangeString(firstSourceCell, lastSourceCell, sourceSheetName)}",
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
					Fields = nameof(CellData.UserEnteredValue).ToCamelCase(),
				},
			};
		}

		/// <summary>
		/// Creates a <see cref="RepeatCellRequest"/> to format a single row of cells (assumes you want to start from the first column)
		/// </summary>
		/// <param name="sheetId">The ID of the sheet</param>
		/// <param name="rowIndex">The index of the row where the formatting should be applied</param>
		/// <param name="endColumnIndex">The index of the column where the formatting should end</param>
		/// <param name="cellDataAction">The action to set the <see cref="CellData"/> properties 
		/// (although technically you could change the values as well, this isn't set up to do so; use an <see cref="AppendRequest"/> or <see cref="UpdateRequest"/> instead)</param>
		/// <param name="propertyNames">A collection of the names of the properties that were set in <paramref name="cellDataAction"/></param>
		/// <returns></returns>
		public static Request CreateRowFormattingRequest(int? sheetId, int rowIndex, int endColumnIndex, Func<CellData> cellDataAction, IEnumerable<string> propertyNames)
		{
			return CreateRowFormattingRequest(sheetId, rowIndex, 0, endColumnIndex, cellDataAction, propertyNames);
		}

		/// <summary>
		/// Creates a <see cref="RepeatCellRequest"/> to format a row of cells
		/// </summary>
		/// <param name="sheetId">The ID of the sheet</param>
		/// <param name="rowIndex">The index of the row where the formatting should be applied</param>
		/// <param name="startColumnIndex">The index of the column where the formatting should start</param>
		/// <param name="cellDataAction">The action to set the <see cref="CellData"/> properties 
		/// (although technically you could change the values as well, this isn't set up to do so; use an <see cref="AppendRequest"/> or <see cref="UpdateRequest"/> instead)</param>
		/// <param name="propertyNames">A collection of the names of the properties that were set in <paramref name="cellDataAction"/></param>
		/// <returns></returns>
		public static Request CreateRowFormattingRequest(int? sheetId, int rowIndex, int startColumnIndex, int endColumnIndex, Func<CellData> cellDataAction, IEnumerable<string> propertyNames)
		{
			return new Request
			{
				RepeatCell = new RepeatCellRequest
				{
					Range = new GridRange
					{
						SheetId = sheetId,
						StartRowIndex = rowIndex,
						StartColumnIndex = startColumnIndex,
						EndRowIndex = rowIndex + 1,
						EndColumnIndex = endColumnIndex + 1,
					},
					Cell = cellDataAction(),
					Fields = propertyNames == null || propertyNames.Count() == 0 ? "*" : $"{nameof(CellData.UserEnteredFormat).ToCamelCase()}({propertyNames.Select(x => x.ToCamelCase()).Aggregate((s1, s2) => $"{s1},{s2}")})",
				},
			};
		}
	}
}
