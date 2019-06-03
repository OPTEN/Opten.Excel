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
		/// <param name="startBody">if set to <c>true</c> start row.</param>
		/// <returns></returns>
		public static DataTable GetDataTableFromExcel(this ExcelPackage package, string worksheet, int startBody)
			=> package.GetDataTableFromExcel(worksheet, null, startBody);

		/// <summary>
		/// Gets the data table from an Excel.
		/// </summary>
		/// <param name="package">The package.</param>
		/// <param name="worksheet">The worksheet.</param>
		/// <param name="startHeader">if set to <c>true</c> start header.</param>
		/// <param name="startBody">if set to <c>true</c> start row.</param>
		/// <returns></returns>
		public static DataTable GetDataTableFromExcel(this ExcelPackage package, string worksheet, int? startHeader, int? startBody)
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
				foreach (ExcelRangeBase firstRowCell in ws.Cells[(startHeader ?? startBody).Value, 1, 1, ws.Dimension.End.Column])
				{
					dt.Columns.Add(startHeader.HasValue ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
				}

				for (int rowNum = startBody.Value; rowNum <= ws.Dimension.End.Row; rowNum++)
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
		public static byte[] WriteDataTableToExcel(this DataTable data, string worksheet)
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				ExcelWorksheet ws = package.Workbook.Worksheets.Add(worksheet);
				ws.Cells["A1"].LoadFromDataTable(data, true);
				return package.GetAsByteArray();
			}
		}

	}
}