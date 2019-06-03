using System.Collections.Generic;
using System.IO;

namespace Opten.Excel.Read
{
	/// <summary>
	/// The CSV Reader.
	/// </summary>
	public abstract class ReadCsvBase : IReader<string>
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
		/// Initializes a new instance of the <see cref="ReadCsvBase"/> class.
		/// </summary>
		/// <param name="path">The path.</param>
		public ReadCsvBase(string path)
		{
			Path = path;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadCsvBase"/> class.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public ReadCsvBase(Stream stream)
		{
			Stream = stream;
		}

		/// <summary>
		/// Reads the CSV.
		/// </summary>
		/// <returns></returns>
		public string[] Read()
		{
			if (this.Stream == null)
			{
				return File.ReadAllLines(path: this.Path);
			}
			else
			{
				List<string> lines = new List<string>();

				using (StreamReader reader = new StreamReader(this.Stream))
				{
					string line;
					while ((line = reader.ReadLine()) != null)
					{
						lines.Add(line);
					}
				}

				return lines.ToArray();
			}
		}
	}
}
