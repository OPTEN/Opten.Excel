using OfficeOpenXml;
using System.Data;
using System.Linq;

namespace Opten.Excel.Extensions
{
	/// <summary>
	/// The DataTable extensions.
	/// </summary>
	public static class DataTableExtensions
	{

		/// <summary>
		/// Gets the data table from an Excel.
		/// </summary>
		/// <param name="package">The package.</param>
		/// <param name="worksheet">The worksheet.</param>
		/// <param name="hasHeader">if set to <c>true</c> [has header].</param>
		/// <returns></returns>
		public static DataTable GetDataTableFromExcel(this ExcelPackage package, string worksheet, bool hasHeader)
		{
			ExcelWorksheet ws = null;

			if (string.IsNullOrWhiteSpace(worksheet))
			{
				ws = package.Workbook.Worksheets.First();
			}
			else
			{
				ws = package.Workbook.Worksheets[worksheet];
			}

			using (DataTable dt = new DataTable())
			{
				foreach (ExcelRangeBase firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
				{
					dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
				}

				int startRow = hasHeader ? 2 : 1;
				for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
				{
					ExcelRangeBase wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];

					// Only rows with text
					if (wsRow.Any() && wsRow.Any(o => string.IsNullOrWhiteSpace(o.Text) == false))
					{
						DataRow row = dt.NewRow();

						foreach (ExcelRangeBase cell in wsRow)
						{
							if (dt.Columns.Count < cell.Start.Column) break;

							row[cell.Start.Column - 1] = cell.Text;
						}

						dt.Rows.Add(row);
					}
				}

				return dt;
			}
		}


		/// <summary>
		/// Writes the data table into the Excel (w/o any styling).
		/// </summary>
		/// <param name="data">The data.</param>
		/// <param name="worksheet">The worksheet.</param>
		/// <returns></returns>
		public static byte[] WriteDataTableIntoExcel(this DataTable data, string worksheet)
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				ExcelWorksheet ws =  package.Workbook.Worksheets.Add(worksheet);
				ws.Cells["A1"].LoadFromDataTable(data, true);
				return package.GetAsByteArray();
			}
		}

	}
}