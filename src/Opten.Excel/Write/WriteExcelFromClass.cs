using Opten.Excel.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Opten.Excel.Write
{
	/// <summary>
	/// The Excel Writer from class.
	/// </summary>
	/// <typeparam name="TClass">The type of the class.</typeparam>
	/// <seealso cref="Opten.Excel.Write.WriteExcelBase" />
	/// <seealso cref="Opten.Excel.Write.IWriter" />
	public class WriteExcelFromClass<TClass> : WriteExcelBase, IWriter
		where TClass : class
	{

		/// <summary>
		/// The mapping columns.
		/// </summary>
		public List<KeyValuePair<string, Func<TClass, string>>> Columns;

		/// <summary>
		/// The row elements.
		/// </summary>
		protected IEnumerable<TClass> Rows;

		/// <summary>
		/// Initializes a new instance of the <see cref="WriteExcelFromClass{TClass}" /> class.
		/// </summary>
		/// <param name="rows">The row elements.</param>
		/// <param name="worksheet">Name of the worksheet.</param>
		/// <exception cref="ArgumentNullException">rowElements;Please provide some elements to write from.</exception>
		public WriteExcelFromClass(IEnumerable<TClass> rows, string worksheet)
			: base(worksheet)
		{
			this.Rows = rows;
		}

		/// <summary>
		/// Converts the data.
		/// </summary>
		/// <param name="printHeaderIfEmpty">if set to <c>true</c> [print header if empty].</param>
		/// <returns></returns>
		/// <exception cref="System.ArgumentNullException">Columns are required to map them from the excel sheet!</exception>
		protected virtual DataTable Convert(bool printHeaderIfEmpty)
		{
			if (Columns == null || Columns.Any() == false)
			{
				throw new ArgumentNullException("Columns are required to map them from the excel sheet!");
			}

			bool hasData = this.Rows != null && this.Rows.Any();

			if (hasData == false && printHeaderIfEmpty == false)
			{
				return null;
			}

			using (DataTable dt = new DataTable())
			{
				// Write header
				foreach (KeyValuePair<string, Func<TClass, string>> column in this.Columns)
				{
					dt.Columns.Add(column.Key);

					if (hasData == false && printHeaderIfEmpty)
					{
						dt.Rows.Add(dt.NewRow());
					}
				}

				// Write rows
				if (hasData)
				{
					foreach (TClass element in this.Rows)
					{
						DataRow row = dt.NewRow();

						for (int i = 0; i < this.Columns.Count; i++)
						{
							KeyValuePair<string, Func<TClass, string>> column = this.Columns[i];

							string result = column.Value.Invoke(element);
							row[i] = result;
						}

						dt.Rows.Add(row);
					}
				}

				return dt;
			}
		}

		/// <summary>
		/// Writes the Excel.
		/// </summary>
		/// <returns></returns>
		public override byte[] Write()
		{
			DataTable data = this.Convert(
				printHeaderIfEmpty: true);

			return data.WriteDataTableToExcel(Worksheet);
		}

	}
}
