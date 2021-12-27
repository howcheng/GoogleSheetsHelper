﻿using System;
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

		private SheetsService Service { get => _service.Value; }
		public Spreadsheet Spreadsheet { get => _spreadsheet.Value; internal set => ResetSpreadsheet(value); }

		public SheetsClient(GoogleCredential credential)
		{
			_service = new Lazy<SheetsService>(() => Init(credential));
			_policy = Policy.Handle<Google.GoogleApiException>()
				.WaitAndRetryAsync(3, (count) =>
				{
					// Google Sheets API has a limit of 100 HTTP requests per 100 seconds per user
					double secondsToWait = 100d - _stopwatch.Elapsed.TotalSeconds;
					Console.WriteLine("Google API request quota reached; waiting {0:00} seconds...", secondsToWait);
					return TimeSpan.FromSeconds(secondsToWait);
				}, (ex, ts) =>
				{
					_stopwatch.Stop();
					_stopwatch.Reset();
					_stopwatch.Start();
				});
		}

		public SheetsClient(GoogleCredential credential, string spreadsheetId)
			: this(credential)
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
			var service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
				HttpClientFactory = _clientFactory,
			});

			return service;
		}

		private Task<Google.Apis.Requests.IDirectResponseSchema> ExecuteRequest(Func<CancellationToken, Task<Google.Apis.Requests.IDirectResponseSchema>> request, CancellationToken ct)
			=> _policy.ExecuteAsync(request, ct);

		private void ResetSpreadsheet(Spreadsheet spreadsheet = null)
		{
			_spreadsheet = new Lazy<Spreadsheet>(() => spreadsheet == null ? GetSpreadsheet().Result : spreadsheet);
		}

		public async Task CreateSpreadsheet(string name, CancellationToken ct = default)
		{
			SpreadsheetsResource resource = new SpreadsheetsResource(Service);
			Spreadsheet spreadsheet = new Spreadsheet();
			spreadsheet.Properties = new SpreadsheetProperties
			{
				Title = name,
			};
			var createRequest = resource.Create(spreadsheet);
			await ExecuteRequest(async token => spreadsheet = await createRequest.ExecuteAsync(token), ct);
			SpreadsheetId = spreadsheet.SpreadsheetId;
			ResetSpreadsheet(spreadsheet);
		}

		private async Task<Spreadsheet> GetSpreadsheet()
		{
			return await Service.Spreadsheets.Get(SpreadsheetId).ExecuteAsync();
		}

		private Sheet GetSheet(string sheetName) => GetSheet(sheetName, Spreadsheet);

		private Sheet GetSheet(string sheetName, Spreadsheet spreadsheet) => spreadsheet.Sheets.FirstOrDefault(s =>
				s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

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

			if (cell.BackgroundColor.HasValue)
				cellData.SetBackgroundColor(cell.BackgroundColor.Value);

			if (cell.ForegroundColor.HasValue)
				cellData.SetForegroundColor(cell.ForegroundColor.Value);

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
			Spreadsheet response = await _policy.ExecuteAsync(async () => await Service.Spreadsheets.Get(SpreadsheetId).ExecuteAsync(ct));
			List<string> result = response.Sheets.Select(x => x.Properties.Title).ToList();
			return result;
		}

		/// <summary>Adds a sheet to the document if it doesn't already exist</summary>
		public async Task<Sheet> GetOrAddSheet(string title, int? columnCount = null, int? rowCount = null, CancellationToken ct = default)
		{
			var sheets = await GetSheetNames(ct);
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
			var sheetId = GetSheetId(sheetName);
			if (sheetId == null)
				throw new ArgumentException($"No such sheet named '{sheetName}'");

			var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>(), IncludeSpreadsheetInResponse = true, };
			var request = new Request
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

		public async Task<Sheet> RenameSheet(string oldName, string newName, CancellationToken ct = default)
		{
			var sheetId = GetSheetId(oldName);
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

			var updateRequest = Service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, Spreadsheet.SpreadsheetId);
			await ExecuteRequest(async token => await updateRequest.ExecuteAsync(token), ct);
			return sheet;
		}

		/// <summary>
		/// Clears the values of a sheet
		/// </summary>
		/// <param name="range">Cell range in A1 notation ("A1:C3") or R1C1 notation (row and column numbers, so "R1C1:R3C3" for "A1:C3") to clear</param>
		/// <param name="ct"></param>
		/// <returns></returns>
		public async Task ClearSheet(string range, CancellationToken ct = default)
		{
			var requestBody = new ClearValuesRequest();
			var deleteRequest = Service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
			_ = await ExecuteRequest(async token => await deleteRequest.ExecuteAsync(token), ct);
		}
		#endregion

		#region Data-related methods
		public async Task<IList<IList<object>>> GetValues(string range, CancellationToken ct = default)
		{
			var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
			request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
			request.DateTimeRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
			var response = await request.ExecuteAsync(ct);
			return response.Values;
		}

		public async Task Append(IList<AppendRequest> data, CancellationToken ct = default)
		{
			var requests = new List<Request>();
			foreach (AppendRequest req in data)
			{
				Request r = CreateAppendRequest(req);
				requests.Add(r);
			}
			await ExecuteRequests(requests, ct);
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
