using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Opten.Excel.Write
{
	/// <summary>
	/// Modifies cells by address of an Excel.
	/// </summary>
	/// <seealso cref="Opten.Excel.Write.IWriter" />
	public class WriteExcelCells : WriteExcelBase, IWriter
	{

		private readonly FileInfo _fileInfo;

		private readonly IDictionary<string, string> _cells;

		/// <summary>
		/// Initializes a new instance of the <see cref="WriteExcelCells" /> class.
		/// </summary>
		/// <param name="fileInfo">The file information.</param>
		/// <param name="cells">The address.</param>
		/// <param name="worksheet">The worksheet.</param>
		/// <exception cref="System.ArgumentNullException">Please provide an excel file to write.;fileInfo
		/// or
		/// Please provide some cells to write.;cells</exception>
		public WriteExcelCells(FileInfo fileInfo, IDictionary<string, string> cells, string worksheet)
			: base(worksheet)
		{
			if (fileInfo == null)
				throw new ArgumentNullException("Please provide an excel file to read/write.", "fileInfo");

			if (cells == null)
				throw new ArgumentNullException("Please provide some cells to write.", "cells");

			_fileInfo = fileInfo;
			_cells = cells;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="WriteExcelCells" /> class.
		/// </summary>
		/// <param name="fileInfo">The file information.</param>
		/// <param name="cells">The address.</param>
		public WriteExcelCells(FileInfo fileInfo, IDictionary<string, string> cells)
			: this(fileInfo, cells, string.Empty)
		{

		}

		/// <summary>
		/// Writes the Excel.
		/// </summary>
		/// <returns></returns>
		public override byte[] Write()
		{
			using (ExcelPackage package = new ExcelPackage(_fileInfo))
			{
				ExcelWorksheet worksheet;
				if (string.IsNullOrWhiteSpace(this.Worksheet))
				{
					worksheet = package.Workbook.Worksheets.First();
				}
				else
				{
					worksheet = package.Workbook.Worksheets[this.Worksheet];
				}

				//TODO: Possibility to address it? -> ["A:B"]
				foreach (KeyValuePair<string, string> address in _cells)
				{
					worksheet.SetValue(address.Key, address.Value);
				}

				package.Save();

				return new byte[0]; //TODO: Why?!
			}
		}

	}
}