using System.Collections.Generic;
using System.Data;

namespace Opten.Excel.Read
{
	/// <summary>
	/// The Excel as CSV Reader.
	/// </summary>
	public abstract class ReadExcelAsCsv : ReadExcelBase<string>, IReader<string>
	{

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelAsCsv"/> class.
		/// </summary>
		/// <param name="path">The path.</param>
		/// <param name="hasHeader">if set to <c>true</c> [has header].</param>
		public ReadExcelAsCsv(string path, bool hasHeader)
			: base(path, hasHeader) { }

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelAsCsv"/> class.
		/// </summary>
		/// <param name="path">The path.</param>
		/// <param name="worksheet">The worksheet.</param>
		/// <param name="hasHeader">if set to <c>true</c> [has header].</param>
		public ReadExcelAsCsv(string path, string worksheet, bool hasHeader)
			: base(path, worksheet, hasHeader) { }

		/// <summary>
		/// Reads the Excel.
		/// </summary>
		/// <returns></returns>
		public override string[] Read()
		{
			DataTable data = GetDataTableFromExcel();

			// Treat Excel like a CSV
			List<string> csv = new List<string>();

			foreach (DataRow row in data.Rows)
			{
				csv.Add(string.Join(";", row.ItemArray));
			}

			return csv.ToArray();
		}

	}
}
