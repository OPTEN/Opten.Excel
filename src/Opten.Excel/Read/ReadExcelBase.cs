using OfficeOpenXml;
using Opten.Excel.Extensions;
using System;
using System.Data;
using System.IO;

namespace Opten.Excel.Read
{
	/// <summary>
	/// The Excel Reader.
	/// </summary>
	/// <typeparam name="TOutput">The type of the output.</typeparam>
	public abstract class ReadExcelBase<TOutput> : IReader<TOutput> where TOutput : class
	{

		/// <summary>
		/// The path.
		/// </summary>
		protected readonly string Path;

		/// <summary>
		/// The stream.
		/// </summary>
		protected readonly Stream Stream;

		/// <summary>
		/// The worksheet.
		/// </summary>
		protected readonly string Worksheet;

		/// <summary>
		/// Dethermines if the worksheet has a header.
		/// </summary>
		protected readonly bool HasHeader;

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelBase{TOutput}"/> class.
		/// </summary>
		/// <param name="path">The path.</param>
		/// <param name="hasHeader">if set to <c>true</c> [has header].</param>
		public ReadExcelBase(string path, bool hasHeader)
		{
			Path = path;
			HasHeader = hasHeader;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelBase{TOutput}"/> class.
		/// </summary>
		/// <param name="stream">The stream.</param>
		/// <param name="hasHeader">if set to <c>true</c> [has header].</param>
		public ReadExcelBase(Stream stream, bool hasHeader)
		{
			Stream = stream;
			HasHeader = hasHeader;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelBase{TOutput}"/> class.
		/// </summary>
		/// <param name="path">The path.</param>
		/// <param name="worksheet">The worksheet.</param>
		/// <param name="hasHeader">if set to <c>true</c> [has header].</param>
		public ReadExcelBase(string path, string worksheet, bool hasHeader)
		{
			Path = path;
			Worksheet = worksheet;
			HasHeader = hasHeader;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelBase{TOutput}"/> class.
		/// </summary>
		/// <param name="stream">The stream.</param>
		/// <param name="worksheet">The worksheet.</param>
		/// <param name="hasHeader">if set to <c>true</c> [has header].</param>
		public ReadExcelBase(Stream stream, string worksheet, bool hasHeader)
		{
			Stream = stream;
			Worksheet = worksheet;
			HasHeader = hasHeader;
		}

		/// <summary>
		/// Reads the Excel.
		/// </summary>
		/// <returns></returns>
		/// <exception cref="System.NotImplementedException"></exception>
		public virtual TOutput[] Read()
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Gets the data table from Excel.
		/// </summary>
		/// <returns></returns>
		protected DataTable GetDataTableFromExcel()
		{
			// Code copied from: http://stackoverflow.com/a/13396787

			if (this.Stream == null)
			{
				using (ExcelPackage package = new ExcelPackage())
				{
					using (FileStream stream = File.OpenRead(this.Path))
					{
						package.Load(stream);
					}

					return package.GetDataTableFromExcel(
						worksheet: this.Worksheet,
						hasHeader: this.HasHeader);
				}
			}
			else
			{
				using (ExcelPackage package = new ExcelPackage(this.Stream))
				{
					return package.GetDataTableFromExcel(
						worksheet: this.Worksheet,
						hasHeader: this.HasHeader);
				}
			}
		}

	}
}