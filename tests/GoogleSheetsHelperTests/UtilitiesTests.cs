using Xunit;

namespace GoogleSheetsHelper.Tests
{
	public class UtilitiesTests
	{
		[Theory]
		[InlineData(0, "A")]
		[InlineData(27, "AB")]
		[InlineData(54, "BC")]
		public void TestConvertIndexToColumnName(int idx, string expected)
		{
			string column = Utilities.ConvertIndexToColumnName(idx);
			Assert.Equal(expected, column);
		}

		[Theory]
		[InlineData(CellRangeOptions.None)]
		[InlineData(CellRangeOptions.FixRow | CellRangeOptions.FixColumn)]
		public void TestCreateCellRangeString(CellRangeOptions options)
		{
			const string START_COL = "A";
			const string END_COL = "B";
			const int START_ROW = 1;
			const int END_ROW = 2;

			string output = Utilities.CreateCellRangeString(START_COL, START_ROW, END_COL, END_ROW, options);
			string expected = string.Format("{0}{1}{0}{2}:{0}{3}{0}{4}", options == CellRangeOptions.None ? string.Empty : "$", START_COL, START_ROW, END_COL, END_ROW);
			Assert.Equal(expected, output);
		}
	}
}
