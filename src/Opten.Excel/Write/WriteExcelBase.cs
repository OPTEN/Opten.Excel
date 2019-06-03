using Opten.Excel.Write;
using System;

namespace Opten.Excel.Write
{
	/// <summary>
	/// The Excel Writer.
	/// </summary>
	public abstract class WriteExcelBase : IWriter
	{

		/// <summary>
		/// The worksheet.
		/// </summary>
		protected readonly string Worksheet;

		/// <summary>
		/// Initializes a new instance of the <see cref="WriteExcelBase"/> class.
		/// </summary>
		/// <param name="worksheet">Name of the worksheet.</param>
		public WriteExcelBase(string worksheet)
		{
			this.Worksheet = worksheet;
		}

		/// <summary>
		/// Writes the Excel.
		/// </summary>
		/// <returns></returns>
		/// <exception cref="System.NotImplementedException"></exception>
		public virtual byte[] Write()
		{
			throw new NotImplementedException();
		}

	}
}
