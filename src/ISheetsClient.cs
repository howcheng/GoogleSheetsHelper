using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	/// <summary>
	/// Interface for classes that manipulate Google Sheets documents
	/// </summary>
	public interface ISheetsClient
	{
		/// <summary>
		/// Gets the spreadsheet ID
		/// </summary>
		string SpreadsheetId { get; }

		/// <summary>
		/// Executes a group of <see cref="Request"/> objects as a batch
		/// </summary>
		/// <param name="requests"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task ExecuteRequests(IEnumerable<Request> requests, CancellationToken ct = default);

		/// <summary>
		/// Creates a new spreadsheet
		/// </summary>
		/// <param name="title"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<Spreadsheet> CreateSpreadsheet(string title, CancellationToken ct = default);
		/// <summary>
		/// Loads an existing spreadsheet
		/// </summary>
		/// <param name="spreadsheetId"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<Spreadsheet> LoadSpreadsheet(string spreadsheetId, CancellationToken ct = default);
		/// <summary>
		/// Renames an existing spreadsheet
		/// </summary>
		/// <param name="newName"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task RenameSpreadsheet(string newName, CancellationToken ct = default);
		/// <summary>
		/// Adds a new sheet to a spreadsheet
		/// </summary>
		/// <param name="title"></param>
		/// <param name="columnCount"></param>
		/// <param name="rowCount"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<Sheet> AddSheet(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default);
		/// <summary>
		/// Gets a sheet by name or adds it if it doesn't exist yet
		/// </summary>
		/// <param name="sheetName"></param>
		/// <param name="columnCount"></param>
		/// <param name="rowCount"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<Sheet> GetOrAddSheet(string sheetName, int? columnCount = null, int? rowCount = null, CancellationToken ct = default);
		/// <summary>
		/// Appends data by adding new cells after the last row with data in a sheet, inserting new rows into the sheet if necessary.
		/// </summary>
		/// <param name="requests"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task Append(IList<AppendRequest> requests, CancellationToken ct = default);
		/// <summary>
		/// Clears all values and formatting from a sheet
		/// </summary>
		/// <param name="sheetName"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task ClearSheet(string sheetName, CancellationToken ct = default);
		/// <summary>
		/// Deletes a sheet
		/// </summary>
		/// <param name="sheetName"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<Spreadsheet> DeleteSheet(string sheetName, CancellationToken ct = default);
		/// <summary>
		/// Renames a sheet
		/// </summary>
		/// <param name="oldName"></param>
		/// <param name="newName"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<Sheet> RenameSheet(string oldName, string newName, CancellationToken ct = default);
		/// <summary>
		/// Gets a list of sheet names
		/// </summary>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<IList<string>> GetSheetNames(CancellationToken ct = default);
		/// <summary>
		/// Gets the row data (values + cell metadata) from a sheet
		/// </summary>
		/// <param name="range">Cell range in A1 ("Sheet1!A1:C3") or R1C1 notation ("Sheet1!R1C1:R3C3"" -- row and column numbers)</param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<IList<RowData>> GetRowData(string range, CancellationToken ct = default);
		/// <summary>
		/// Gets multiple sets of row data (values + cell metadata) from a document
		/// </summary>
		/// <param name="ranges">Cell ranges in A1 ("Sheet1!A1:C3") or R1C1 notation ("Sheet1!R1C1:R3C3"" -- row and column numbers)</param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task<IList<IList<RowData>>> GetRowData(IEnumerable<string> ranges, CancellationToken ct = default);
		/// <summary>
		/// Gets the values from a sheet; NOTE: null values will be skipped!
		/// </summary>
		/// <param name="range">Cell range in A1 ("Sheet1!A1:C3") or R1C1 notation ("Sheet1!R1C1:R3C3"" -- row and column numbers)</param>
		/// <param name="ct"></param>
		/// <returns></returns>
		/// <remarks>It's Google's code that makes the choice to skip null values, so we can't change that; see <see cref="ValueRange.Values"/>.</remarks>
		Task<IList<IList<object>>> GetValues(string range, CancellationToken ct = default);
		/// <summary>
		/// Updates values in a sheet
		/// </summary>
		/// <param name="data"></param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task Update(IList<UpdateRequest> data, CancellationToken ct = default);
		/// <summary>
		/// Clears the values of a sheet (but not the formatting)
		/// </summary>
		/// <param name="range">Cell range in A1 notation ("A1:C3") or R1C1 notation (row and column numbers, so "R1C1:R3C3" for "A1:C3") to clear<. For an entire sheet, use the sheet name by itself./param>
		/// <param name="ct"></param>
		/// <returns></returns>
		Task ClearValues(string range, CancellationToken ct = default);
		/// <summary>
		/// Resizes a column to be the width of its longest value
		/// </summary>
		/// <param name="sheetName"></param>
		/// <param name="columnIndex"></param>
		/// <returns></returns>
		Task<int> AutoResizeColumn(string sheetName, int columnIndex);
	}
}