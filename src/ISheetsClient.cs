using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	public interface ISheetsClient
	{
		string SpreadsheetId { get; }

		Task<Spreadsheet> AddSheet(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default);
		Task<Spreadsheet> AddSheetIfNotExists(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default);
		Task Append(IList<AppendRequest> data, CancellationToken ct = default);
		Task ClearSheet(string range, CancellationToken ct = default);
		Task<Spreadsheet> DeleteSheet(string sheetName, CancellationToken ct = default);
		Task<IList<IList<object>>> GetOrAddSheet(string range, CancellationToken ct = default);
		Task<IList<string>> GetSheets(CancellationToken ct = default);
		Task<IList<IList<object>>> GetValues(string range, CancellationToken ct = default);
		Task Update(IList<UpdateRequest> data, CancellationToken ct = default);
	}
}