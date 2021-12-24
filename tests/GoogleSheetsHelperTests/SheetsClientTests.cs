using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using AutoFixture;
using Google.Apis.Http;
using Google.Apis.Sheets.v4.Data;
using Moq;
using Moq.Protected;
using Xunit;

namespace GoogleSheetsHelper.Tests
{
	public class SheetsClientTests
	{
		private const string SPREADSHEET_ID = "abcdefghijklmnopqrstuvwxyz";
		private const string SPREADSHEET_TITLE = "SPREADSHEET TITLE";
		private const int SHEET_ID = 1234;

		private HttpResponseMessage CreateResponse(string json)
		{
			HttpResponseMessage response = new HttpResponseMessage();
			response.Content = new StringContent(json, System.Text.Encoding.UTF8, System.Net.Mime.MediaTypeNames.Application.Json);
			return response;
		}

		private SheetsClient GetClient(HttpResponseMessage response)
		{
			// the Google Sheets client classes don't implement interfaces, but we can mock the JSON response from Google

			Mock<HttpMessageHandler> mockHandler = new Mock<HttpMessageHandler>();
			mockHandler.Protected().Setup<Task<HttpResponseMessage>>(
				"SendAsync",
				ItExpr.IsAny<HttpRequestMessage>(),
				ItExpr.IsAny<CancellationToken>()).ReturnsAsync(response);

			ConfigurableHttpClient client = new ConfigurableHttpClient(new ConfigurableMessageHandler(mockHandler.Object));

			Mock<IHttpClientFactory> mockFactory = new Mock<IHttpClientFactory>();
			mockFactory.Setup(x => x.CreateHttpClient(It.IsAny<CreateHttpClientArgs>())).Returns(client);
			
			return new SheetsClient(mockFactory.Object, SPREADSHEET_ID);
		}

		[Fact]
		public async Task TestCreateSpreadsheet()
		{
			// JSON from https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#Spreadsheet
			Fixture f = new Fixture();
			string id = f.Create<string>();
			string json = $@"{{
	""spreadsheetId"": ""{id}"",
	""properties"": {{ ""title"": ""{SPREADSHEET_TITLE}"" }},
	""sheets"": [
		{{ 
			""properties"": {{ ""sheetId"": {SHEET_ID}, ""title"": ""Sheet 1"", ""index"": 0 }}
		}}
	]
}}";

			SheetsClient client = GetClient(CreateResponse(json));
			await client.CreateSpreadsheet(SPREADSHEET_TITLE);
			Assert.Equal(id, client.SpreadsheetId);
		}

		[Fact]
		public async Task TestExecuteRequests()
		{
			Fixture f = new Fixture();
			IEnumerable<Request> requests = f.CreateMany<Request>();

			// JSON from https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate
			string json = $@"{{
	""spreadsheetId"": ""{SPREADSHEET_ID}"",
	""updatedSpreadsheet"": {{
		""spreadsheetId"": ""{SPREADSHEET_ID}"",
		""properties"": {{ ""title"": ""{SPREADSHEET_TITLE}"" }},
		""sheets"": [
			{{ 
				""properties"": {{ ""sheetId"": {SHEET_ID}, ""title"": ""Sheet 1"", ""index"": 0 }}
			}}
		]
	}}
}}";
			SheetsClient client = GetClient(CreateResponse(json));
			await client.ExecuteRequests(requests);
			Assert.NotNull(client.Spreadsheet); // since the Spreadsheet is null upon initialization, this should be populated now with our mocked document
		}

		[Fact]
		public async Task TestGetSheetNames()
		{
			Fixture f = new Fixture();
			List<string> sheetNames = f.CreateMany<string>(3).ToList();

			string json = $@"{{
	""spreadsheetId"": ""{SPREADSHEET_ID}"",
	""properties"": {{ ""title"": ""{SPREADSHEET_TITLE}"" }},
	""sheets"": [
		{{ 
			""properties"": {{ ""sheetId"": {SHEET_ID}, ""title"": ""{sheetNames.ElementAt(0)}"", ""index"": 0 }}
		}},
		{{
			""properties"": {{ ""sheetId"": 1235, ""title"": ""{sheetNames.ElementAt(1)}"", ""index"": 1 }}
		}},
		{{
			""properties"": {{ ""sheetId"": 1236, ""title"": ""{sheetNames.ElementAt(2)}"", ""index"": 2 }}
		}}
	]
}}";

			SheetsClient client = GetClient(CreateResponse(json));
			IEnumerable<string> names = await client.GetSheetNames();
			Assert.Equal(sheetNames, names);
		}

		[Fact]
		public async Task TestAddSheet()
		{
			const string SHEET_NAME = "SHEET NAME";
			string json = $@"{{
	""spreadsheetId"": ""{SPREADSHEET_ID}"",
	""updatedSpreadsheet"": {{
		""spreadsheetId"": ""{SPREADSHEET_ID}"",
		""properties"": {{ ""title"": ""{SPREADSHEET_TITLE}"" }},
		""sheets"": [
			{{ 
				""properties"": {{ ""sheetId"": {SHEET_ID}, ""title"": ""{SHEET_NAME}"", ""index"": 0 }}
			}}
		]
	}}
}}";

			SheetsClient client = GetClient(CreateResponse(json));
			Sheet sheet = await client.AddSheet(SHEET_NAME, 3, 5);
			Assert.Equal(SHEET_NAME, sheet.Properties.Title);
		}

		[Fact]
		public async Task TestDeleteNonexistentSheet()
		{
			Spreadsheet spreadsheet = new Spreadsheet
			{
				SpreadsheetId = SPREADSHEET_ID,
				Sheets = new List<Sheet>
				{
					new Sheet
					{
						Properties = new SheetProperties
						{
							Title = "SHEET NAME",
						}
					}
				}
			};

			SheetsClient client = GetClient(new HttpResponseMessage());
			client.Spreadsheet = spreadsheet;
			await Assert.ThrowsAsync<ArgumentException>(async () => await client.DeleteSheet("NOT THE NAME"));
		}

		[Fact]
		public async Task TestAutoResizeColumn()
		{
			const string SHEET_NAME = "SHEET NAME";
			const int PIXEL_SIZE = 100;
			Spreadsheet spreadsheet = new Spreadsheet
			{
				SpreadsheetId = SPREADSHEET_ID,
				Sheets = new List<Sheet>
				{
					new Sheet
					{
						Properties = new SheetProperties
						{
							Title = SHEET_NAME,
							SheetId = SHEET_ID,
						}
					}
				}
			};

			// https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#GridData
			// https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#DimensionProperties
			string json = $@"{{
	""spreadsheetId"": ""{SPREADSHEET_ID}"",
	""updatedSpreadsheet"": {{
		""spreadsheetId"": ""{SPREADSHEET_ID}"",
		""properties"": {{ ""title"": ""{SPREADSHEET_TITLE}"" }},
		""sheets"": [
			{{ 
				""properties"": {{ ""sheetId"": {SHEET_ID}, ""title"": ""{SHEET_NAME}"", ""index"": 0 }},
				""data"": [
					{{ ""columnMetadata"": [ {{ ""pixelSize"": {PIXEL_SIZE} }} ] }}
				]
			}}
		]
	}}
}}";
			SheetsClient client = GetClient(CreateResponse(json));
			client.Spreadsheet = spreadsheet;
			int newValue = await client.AutoResizeColumn(SHEET_NAME, 1);

			Assert.Equal(PIXEL_SIZE, newValue);
		}
	}
}