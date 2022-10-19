using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Logging;
using Polly;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("GoogleSheetsHelper.Tests")]
namespace GoogleSheetsHelper
{
	public class SheetsClient : ISheetsClient
	{
		private static readonly string ApplicationName = "GoogleSheetsHelper";

		public string SpreadsheetId { get; private set; }

		private Lazy<SheetsService> _service;
		private Lazy<Spreadsheet> _spreadsheet;
		private static Stopwatch _stopwatch = new Stopwatch();
		private AsyncPolicy _policy;
		private IHttpClientFactory _clientFactory;
		private ILogger<SheetsClient> _log;

		private SheetsService Service { get => _service.Value; }
		public Spreadsheet Spreadsheet { get => _spreadsheet.Value; internal set => ResetSpreadsheet(value); }

		public SheetsClient(GoogleCredential credential, ILogger<SheetsClient> log)
		{
			_service = new Lazy<SheetsService>(() => Init(credential));
			_log = log;
			_policy = Policy.Handle<Google.GoogleApiException>()
				.WaitAndRetryAsync(3, (count) =>
				{
					// Google Sheets API has a limit of 100 HTTP requests per 100 seconds per user
					double secondsToWait = 100d - _stopwatch.Elapsed.TotalSeconds;
					_log.LogInformation("Google API request quota reached; waiting {0:00} seconds...", secondsToWait);
					return TimeSpan.FromSeconds(secondsToWait);
				}, (ex, ts) =>
				{
					_stopwatch.Stop();
					_stopwatch.Reset();
					_stopwatch.Start();
				});
		}

		public SheetsClient(GoogleCredential credential, string spreadsheetId, ILogger<SheetsClient> log)
			: this(credential, log)
		{
			SpreadsheetId = spreadsheetId;
			ResetSpreadsheet();
		}

		/// <summary>
		/// This constructor is for unit testing
		/// </summary>
		/// <param name="clientFactory"></param>
		internal SheetsClient(IHttpClientFactory clientFactory, string spreadsheetId)
		{
			_clientFactory = clientFactory;
			_service = new Lazy<SheetsService>(() => Init(null));
			_policy = Policy.Handle<Google.GoogleApiException>().RetryAsync(1);
			SpreadsheetId = spreadsheetId;
		}

		private SheetsService Init(GoogleCredential credential)
		{
			BaseClientService.Initializer initializer = new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			};
			if (_clientFactory != null)
				initializer.HttpClientFactory = _clientFactory;

			return new SheetsService(initializer);
		}

		private async Task<Google.Apis.Requests.IDirectResponseSchema> ExecuteRequest(Func<CancellationToken, Task<Google.Apis.Requests.IDirectResponseSchema>> request, CancellationToken ct)
		{
			_stopwatch.Start();
			Google.Apis.Requests.IDirectResponseSchema response = await _policy.ExecuteAsync(request, ct);
			_stopwatch.Stop();
			_stopwatch.Reset();
			return response;
		}

		private void ResetSpreadsheet(Spreadsheet spreadsheet = null)
		{
			_spreadsheet = new Lazy<Spreadsheet>(() => spreadsheet == null ? GetSpreadsheet().Result : spreadsheet);
		}

		public async Task<Spreadsheet> CreateSpreadsheet(string name, CancellationToken ct = default)
		{
			SpreadsheetsResource resource = new SpreadsheetsResource(Service);
			Spreadsheet spreadsheet = new Spreadsheet();
			spreadsheet.Properties = new SpreadsheetProperties
			{
				Title = name,
			};
			SpreadsheetsResource.CreateRequest createRequest = resource.Create(spreadsheet);
			spreadsheet = (Spreadsheet)await ExecuteRequest(async token => await createRequest.ExecuteAsync(token), ct);
			SpreadsheetId = spreadsheet.SpreadsheetId;
			ResetSpreadsheet(spreadsheet);
			return spreadsheet;
		}

		public async Task<Spreadsheet> LoadSpreadsheet(string spreadsheetId, CancellationToken ct = default)
		{
			SpreadsheetId = spreadsheetId;
			Spreadsheet spreadsheet = await GetSpreadsheet(ct);
			ResetSpreadsheet(spreadsheet);
			return spreadsheet;
		}

		private async Task<Spreadsheet> GetSpreadsheet(CancellationToken ct = default)
		{
			SpreadsheetsResource.GetRequest getRequest = Service.Spreadsheets.Get(SpreadsheetId);
			return (Spreadsheet)await ExecuteRequest(async token => await getRequest.ExecuteAsync(token), ct);
		}

		public async Task RenameSpreadsheet(string newName, CancellationToken ct = default)
		{
			Request request = new Request
			{
				UpdateSpreadsheetProperties = new UpdateSpreadsheetPropertiesRequest
				{
					Properties = new SpreadsheetProperties
					{
						Title = newName
					},
					Fields = nameof(SpreadsheetProperties.Title).ToCamelCase()
				}
			};
			await ExecuteRequests(new[] { request }, ct);
		}

		private Sheet GetSheet(string sheetName) => GetSheet(sheetName, Spreadsheet);

		private Sheet GetSheet(string sheetName, Spreadsheet spreadsheet) 
			=> spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

		private int? GetSheetId(string sheetName)
		{
			Sheet sheet = GetSheet(sheetName);
			return sheet?.Properties.SheetId;
		}

		private CellData CreateCellData(GoogleSheetCell cell)
		{
			if (cell == null)
				return new CellData();

			var numberValue = cell.NumberValue;

			if (cell.DateTimeValue.HasValue)
			{
				numberValue = cell.DateTimeValue.Value.ToOADate();
				if (string.IsNullOrWhiteSpace(cell.DateTimeFormat))
					cell.DateTimeFormat = "yyyy-mm-dd";
			}

			var userEnteredValue = new ExtendedValue
			{
				StringValue = cell.StringValue,
				NumberValue = numberValue,
				BoolValue = cell.BoolValue,
				FormulaValue = cell.FormulaValue,
			};

			CellData cellData = new CellData
			{
				UserEnteredValue = userEnteredValue,
			};

			if (cell.NumberValue.HasValue && !string.IsNullOrEmpty(cell.NumberFormat))
				cellData.SetNumberFormat(cell.NumberFormat);

			if (cell.DateTimeValue.HasValue && !string.IsNullOrEmpty(cell.DateTimeFormat))
				cellData.SetNumberFormat(cell.DateTimeFormat);

			if (cell.Bold.HasValue)
				cellData.SetBoldText(cell.Bold.Value);

			cellData.SetBackgroundColor(cell.GoogleBackgroundColor);
			cellData.SetForegroundColor(cell.GoogleForegroundColor);

			if (cell.HorizontalAlignment.HasValue)
				cellData.SetHorizontalAlignment(cell.HorizontalAlignment.Value);

			return cellData;
		}

		public async Task ExecuteRequests(IEnumerable<Request> requests, CancellationToken ct = default)
		{
			if (requests.Count() == 0)
				return;

			Spreadsheet spreadsheet = await ExecuteRequests(requests, null, ct);
			ResetSpreadsheet(spreadsheet);
		}

		// this overload does NOT reset the Spreadsheet property!
		private async Task<Spreadsheet> ExecuteRequests(IEnumerable<Request> requests, Action<BatchUpdateSpreadsheetRequest> configureAction, CancellationToken ct = default)
		{
			if (requests.Count() == 0)
				return Spreadsheet;

			SpreadsheetsResource resource = new SpreadsheetsResource(Service);
			BatchUpdateSpreadsheetRequest updateRequest = new BatchUpdateSpreadsheetRequest
			{
				Requests = requests.ToList(),
				IncludeSpreadsheetInResponse = true,
			};
			if (configureAction != null)
				configureAction(updateRequest);

			SpreadsheetsResource.BatchUpdateRequest batchUpdate = resource.BatchUpdate(updateRequest, SpreadsheetId);
			var batchUpdateResponse = (BatchUpdateSpreadsheetResponse)await ExecuteRequest(async token => await batchUpdate.ExecuteAsync(token), ct);
			return batchUpdateResponse.UpdatedSpreadsheet;
		}

		#region Sheet-related methods
		/// <summary>Gets a list of sheet names</summary>
		public async Task<IList<string>> GetSheetNames(CancellationToken ct = default)
		{
			Spreadsheet response = await GetSpreadsheet(ct);
			List<string> result = response.Sheets.Select(x => x.Properties.Title).ToList();
			return result;
		}

		/// <summary>Adds a sheet to the document if it doesn't already exist</summary>
		public async Task<Sheet> GetOrAddSheet(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default)
		{
			IList<string> sheets = await GetSheetNames(ct);
			if (sheets.Any(x => string.Equals(x, title, StringComparison.OrdinalIgnoreCase)))
				return Spreadsheet.Sheets.First(x => string.Equals(x.Properties.Title, title, StringComparison.OrdinalIgnoreCase));
			
			return await AddSheet(title, columnCount, rowCount, ct);
		}

		/// <summary>
		/// Adds a new sheet
		/// </summary>
		/// <param name="title">Name of the sheet</param>
		/// <param name="columnCount">Number of columns</param>
		/// <param name="rowCount">Number of rows</param>
		public async Task<Sheet> AddSheet(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default)
		{
			var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
			var request = new Request
			{
				AddSheet = new AddSheetRequest
				{
					Properties = new SheetProperties
					{
						Title = title,
						GridProperties = new GridProperties
						{
							ColumnCount = columnCount,
							RowCount = rowCount,
						}
					}
				}
			};
			requests.Requests.Add(request);
			var response = (BatchUpdateSpreadsheetResponse)await ExecuteRequest(async token => await Service.Spreadsheets.BatchUpdate(requests, SpreadsheetId).ExecuteAsync(token), ct);
			ResetSpreadsheet(response.UpdatedSpreadsheet);
			return Spreadsheet.Sheets.Single(x => x.Properties.Title == title);
		}

		public async Task<Spreadsheet> DeleteSheet(string sheetName, CancellationToken ct = default)
		{
			int? sheetId = GetSheetId(sheetName);
			if (sheetId == null)
				throw new ArgumentException($"No such sheet named '{sheetName}'");

			var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>(), IncludeSpreadsheetInResponse = true, };
			Request request = new Request
			{
				DeleteSheet = new DeleteSheetRequest
				{
					SheetId = sheetId
				}
			};
			requests.Requests.Add(request);
			BatchUpdateSpreadsheetResponse response = (BatchUpdateSpreadsheetResponse)await ExecuteRequest(async token => await Service.Spreadsheets.BatchUpdate(requests, SpreadsheetId).ExecuteAsync(token), ct);
			ResetSpreadsheet(response.UpdatedSpreadsheet);
			return Spreadsheet;
		}

		public async Task ClearSheet(string sheetName, CancellationToken ct = default)
		{
			int? sheetId = GetSheetId(sheetName);
			List<Task> tasks = new List<Task>();
			tasks.Add(ClearValues($"'{sheetName}'", ct));

			// https://stackoverflow.com/questions/45801313/remove-only-formatting-on-a-cell-range-selection-with-google-spreadsheet-api
			Request request = new Request
			{
				UpdateCells = new UpdateCellsRequest
				{
					Range = new GridRange
					{
						SheetId = sheetId,
					},
					Fields = nameof(CellData.UserEnteredFormat).ToCamelCase(),
				},
			};
			tasks.Add(ExecuteRequests(new[] { request }, ct));
			await Task.WhenAll(tasks);
		}

		public async Task<Sheet> RenameSheet(string oldName, string newName, CancellationToken ct = default)
		{
			int? sheetId = GetSheetId(oldName);
			if (sheetId == null)
				throw new ArgumentException($"No such sheet named '{oldName}'");

			Sheet sheet = Spreadsheet.Sheets.Single(x => x.Properties.SheetId == sheetId);
			sheet.Properties.Title = newName;

			List<Request> requests = new List<Request>();
			requests.Add(new Request
			{
				UpdateSheetProperties = new UpdateSheetPropertiesRequest
				{
					Properties = sheet.Properties,
					Fields = nameof(SheetProperties.Title).ToLower(),
				},
			});

			SpreadsheetsResource.BatchUpdateRequest updateRequest = Service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, Spreadsheet.SpreadsheetId);
			await ExecuteRequest(async token => await updateRequest.ExecuteAsync(token), ct);
			return sheet;
		}
		#endregion

		#region Data-related methods
		public async Task<IList<RowData>> GetRowData(string range, CancellationToken ct = default)
			=> (await GetRowData(new[] { range })).FirstOrDefault();

		public async Task<IList<IList<RowData>>> GetRowData(IEnumerable<string> ranges, CancellationToken ct = default)
		{
			SpreadsheetsResource.GetRequest request = Service.Spreadsheets.Get(SpreadsheetId);
			request.Ranges = ranges.ToList();
			request.IncludeGridData = true;
			_log.LogDebug($"{nameof(GetRowData)}: Getting RowData objects for ranges: {ranges.Aggregate((s1, s2) => $"{s1}, {s2}")}");
			Spreadsheet spreadsheet = (Spreadsheet)await ExecuteRequest(async token => await request.ExecuteAsync(token), ct);

			List<IList<RowData>> ret = new List<IList<RowData>>(spreadsheet.Sheets.Count);
			foreach (Sheet sheet in spreadsheet.Sheets)
			{
				if (sheet.Data == null && sheet.Data.Count == 0)
					continue;

				foreach (GridData gridData in sheet.Data)
				{
					ret.Add(gridData.RowData);
				}
			}
			return ret;
		}

		public async Task<IList<IList<object>>> GetValues(string range, CancellationToken ct = default)
		{
			SpreadsheetsResource.ValuesResource.GetRequest request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
			request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
			request.DateTimeRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
			ValueRange response = (ValueRange)await ExecuteRequest(async token => await request.ExecuteAsync(token), ct);
			return response.Values;
		}

		public async Task Append(IList<AppendRequest> request, CancellationToken ct = default)
		{
			foreach (AppendRequest req in request)
			{
				// we need to do these serially because when done in batch the requests are done in parallel
				// and because Append just puts data at the end, because of a race conditions, some actions will overwrite others
				Request r = CreateAppendRequest(req);
				await ExecuteRequests(new[] { r }, ct);
			}
		}

		private Request CreateAppendRequest(AppendRequest r)
		{
			var sheetId = GetSheetId(r.SheetName);
			if (sheetId == null)
				throw new ArgumentException($"No such sheet named '{r.SheetName}'");

			var listRowData = new List<RowData>();
			var request = new Request
			{
				AppendCells = new AppendCellsRequest
				{
					SheetId = sheetId,
					Rows = listRowData,
					Fields = "*",
				},
			};

			foreach (var row in r.Rows)
			{
				var listCellData = new List<CellData>();

				foreach (var cell in row)
				{
					var cellData = CreateCellData(cell);
					listCellData.Add(cellData);
				}

				var rowData = new RowData() { Values = listCellData };
				listRowData.Add(rowData);
			}

			return request;
		}

		/// <summary>Updates data in a Google sheet</summary>
		public async Task Update(IList<UpdateRequest> data, CancellationToken ct = default)
		{
			var requests = new List<Request>();
			foreach (UpdateRequest req in data)
			{
				Request r = CreateUpdateRequest(req);
				requests.Add(r);
			}
			await ExecuteRequests(requests, ct);
		}

		private Request CreateUpdateRequest(UpdateRequest r)
		{
			var sheetId = GetSheetId(r.SheetName);
			if (sheetId == null)
				throw new ArgumentException($"No such sheet named '{r.SheetName}'");

			var gc = new GridCoordinate
			{
				ColumnIndex = r.ColumnStart,
				RowIndex = r.RowStart,
				SheetId = sheetId
			};
			var request = new Request
			{
				UpdateCells = new UpdateCellsRequest { Start = gc, Fields = "*" }
			};
			var listRowData = new List<RowData>();

			foreach (var row in r.Rows)
			{
				var listCellData = new List<CellData>();

				foreach (var cell in row)
				{
					var cellData = CreateCellData(cell);
					listCellData.Add(cellData);
				}

				var rowData = new RowData() { Values = listCellData };
				listRowData.Add(rowData);
			}

			request.UpdateCells.Rows = listRowData;
			return request;
		}

		public async Task UpdateValues(IList<UpdateRequest> requests, CancellationToken ct = default)
		{
			foreach (UpdateRequest updateRequest in requests)
			{
				int startRowNum = updateRequest.RowStart + 1;
				int endColumnIdx = updateRequest.ColumnStart + updateRequest.Rows.Max(x => x.Count);
				int endRowNum = startRowNum + updateRequest.Rows.Count;
				string startCell = $"{Utilities.ConvertIndexToColumnName(updateRequest.ColumnStart)}{startRowNum}";
				string endCell = $"{Utilities.ConvertIndexToColumnName(endColumnIdx)}{endRowNum}";
				string range = $"'{updateRequest.SheetName}'!{startCell}:{endCell}";
				ValueRange valueRange = new ValueRange
				{
					Range = range,
					MajorDimension = "ROWS",
					Values = updateRequest.Rows.Select(r => (IList<object>)r.Select(c => c.CellValue).ToList()).ToList(),
				};
				SpreadsheetsResource.ValuesResource.UpdateRequest request = Service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
				request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
				await ExecuteRequest(async token => await request.ExecuteAsync(token), ct);
			}
		}

		/// <summary>
		/// Clears the values of a sheet
		/// </summary>
		/// <param name="range">Cell range in A1 notation ("A1:C3") or R1C1 notation (row and column numbers, so "R1C1:R3C3" for "A1:C3") to clear</param>
		/// <param name="ct"></param>
		/// <returns></returns>
		public async Task ClearValues(string range, CancellationToken ct = default)
		{
			ClearValuesRequest requestBody = new ClearValuesRequest();
			SpreadsheetsResource.ValuesResource.ClearRequest valuesRequest = Service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
			_ = await ExecuteRequest(async token => await valuesRequest.ExecuteAsync(token), ct);
		}
		#endregion

		#region Column-related methods
		/// <summary>
		/// Sends a request to auto-resize a column
		/// </summary>
		/// <param name="sheetName">The name of the sheet where the column is</param>
		/// <param name="columnIndex">Index of the column to be resized</param>
		public async Task<int> AutoResizeColumn(string sheetName, int columnIndex)
		{
			Sheet sheet = GetSheet(sheetName);
			Request resizeRequest = new Request
			{
				AutoResizeDimensions = new AutoResizeDimensionsRequest
				{
					Dimensions = new DimensionRange
					{
						Dimension = SpreadsheetsResource.ValuesResource.BatchGetRequest.MajorDimensionEnum.COLUMNS.ToString(),
						SheetId = sheet.Properties.SheetId,
						StartIndex = columnIndex,
						EndIndex = columnIndex + 1,
					},
				},
			};
			// get the new width value in the response
			string columnName = Utilities.ConvertIndexToColumnName(columnIndex);
			Spreadsheet updatedSheet = await ExecuteRequests(new[] { resizeRequest }, rq => 
			{
				rq.ResponseIncludeGridData = true;
				rq.ResponseRanges = new List<string> { $"{sheet.Properties.Title}!{columnName}1" };
			});

			sheet = GetSheet(sheetName, updatedSheet);
			return sheet.Data.First().ColumnMetadata.First().PixelSize.Value; // we are only getting a range of one cell in the response
		}
		#endregion

		#region Value parsing methods
		public static bool TryParseDateTime(object obj, out DateTime result)
		{
			result = default;
			try
			{
				if (obj == null)
				{
					return false;
				}

				if (obj is double d)
				{
					result = DateTime.FromOADate(d);
					return true;
				}

				if (obj is long l)
				{
					result = DateTime.FromOADate(l);
					return true;
				}

				if (obj is int i)
				{
					result = DateTime.FromOADate(i);
					return true;
				}

				if (obj is float f)
				{
					result = DateTime.FromOADate(f);
					return true;
				}

				if (double.TryParse(obj.ToString(), out var dt))
				{
					result = DateTime.FromOADate(dt);
					return true;
				}
			}
			catch
			{
			}
			return false;
		}

		public static bool TryParseDouble(object s, out double result)
		{
			try
			{
				result = Convert.ToDouble(s);
				return true;
			}
			catch
			{
				result = default;
				return false;
			}
		}
		#endregion
	}
}
