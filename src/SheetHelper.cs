using System;
using System.Collections.Generic;
using System.Linq;

namespace GoogleSheetsHelper
{
	/// <summary>
	/// Base helper class to get column indices or names based on the contents of a header row
	/// </summary>
	public class SheetHelper
	{
		/// <summary>
		/// A collection of all the header columns
		/// </summary>
		public List<string> HeaderRowColumns { get; private set; }

		public SheetHelper(IEnumerable<string> headerRowColumns)
		{
			HeaderRowColumns = headerRowColumns.ToList();
		}

		/// <summary>
		/// Gets the column index (zero-based) by header value
		/// </summary>
		/// <param name="colHeader"></param>
		/// <returns>The zero-based index of the column header in <see cref="HeaderRowColumns"/>, or <c>-1</c> if it doesn't exist.</returns>
		public virtual int GetColumnIndexByHeader(string colHeader)
		{
			int idx = HeaderRowColumns.IndexOf(colHeader);
			return idx;
		}

		/// <summary>
		/// Gets the Google Sheets column name by header value.
		/// </summary>
		/// <param name="colHeader"></param>
		/// <returns>The Google Sheets column name, e.g. "A" for the first item in <see cref="HeaderRowColumns"/, or <c>null</c> if it doesn't exist.</returns>
		public string GetColumnNameByHeader(string colHeader)
		{
			byte idx = (byte)GetColumnIndexByHeader(colHeader);
			return Utilities.ConvertIndexToColumnName(idx);
		}

		/// <summary>
		/// Creates a <see cref="GoogleSheetRow"/> with the given header values
		/// </summary>
		/// <param name="headers"></param>
		/// <param name="formatter"></param>
		/// <returns></returns>
		public GoogleSheetRow CreateHeaderRow(IEnumerable<string> headers, Action<GoogleSheetCell> formatter = null)
		{
			GoogleSheetRow row = new GoogleSheetRow();
			row.AddRange(CreateHeaderCells(headers, formatter));
			return row;
		}

		/// <summary>
		/// Creates a collection of <see cref="GoogleSheetCell"/> objects from the given header values
		/// </summary>
		/// <param name="headers"></param>
		/// <param name="formatter"></param>
		/// <returns></returns>
		public IEnumerable<GoogleSheetCell> CreateHeaderCells(IEnumerable<string> headers, Action<GoogleSheetCell> formatter = null)
		{
			return headers.Select(x =>
			{
				GoogleSheetCell cell = new GoogleSheetCell(x);
				formatter?.Invoke(cell);
				return cell;
			});
		}
	}
}
