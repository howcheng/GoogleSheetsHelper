using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;
using Xunit;

namespace GoogleSheetsHelper.Tests
{
	public class RequestCreatorTests
	{
		[Fact]
		public void TestPropertyNamesAreFormattedCorrectly()
		{
			Request? request = RequestCreator.CreateRowFormattingRequest(123, 0, 0, 10, () =>
				new CellData
				{
					UserEnteredFormat = new CellFormat
					{
						BackgroundColor = System.Drawing.Color.White.ToGoogleColor(),
						TextFormat = new TextFormat
						{
							Bold = true,
						}
					}
				},
				new[] { nameof(CellFormat.BackgroundColor), nameof(CellFormat.TextFormat) }
			);

			Assert.Equal("userEnteredFormat(backgroundColor,textFormat)", request.RepeatCell.Fields);
		}
	}
}
