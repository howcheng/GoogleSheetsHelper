using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsHelper
{
	public class AppendRequest
	{
		public string SheetName { get; set; }
		public IList<GoogleSheetRow> Rows { get; set; }

		public AppendRequest(string sheetName)
		{
			SheetName = sheetName;
			Rows = new List<GoogleSheetRow>();
		}
	}
}
