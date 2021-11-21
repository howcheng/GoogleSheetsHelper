using System.Collections.Generic;

namespace GoogleSheetsHelper
{
	/// <summary>
	/// Represents a request to update a Google Sheet
	/// </summary>
	public class UpdateRequest
	{
		public string SheetName { get; set; }
		/// <summary>
		/// The starting column index
		/// </summary>
		public int ColumnStart { get; set; }
		/// <summary>
		/// The starting row index
		/// </summary>
		public int RowStart { get; set; }
		public IList<GoogleSheetRow> Rows { get; set; }

		public UpdateRequest(string sheetName)
		{
			SheetName = sheetName;
			Rows = new List<GoogleSheetRow>();
		}
	}
}
