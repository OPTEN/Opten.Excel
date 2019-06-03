using Opten.Core.Extensions;
using Opten.Core.Parsers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace Opten.Excel.Read
{
	/// <summary>
	/// The Excel Reader from class.
	/// </summary>
	/// <typeparam name="TClass">The type of the class.</typeparam>
	public abstract class ReadExcelFromClass<TClass> : ReadExcelBase<TClass>, IReader<TClass> where TClass : class
	{

		/// <summary>
		/// The mapping columns.
		/// </summary>
		protected Dictionary<string[], Expression<Func<TClass, dynamic>>> Columns;

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelFromClass{TClass}"/> class.
		/// </summary>
		/// <param name="path">The path.</param>
		public ReadExcelFromClass(string path)
			: base(path, true) { }

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelFromClass{TClass}"/> class.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public ReadExcelFromClass(Stream stream)
			: base(stream, true) { }

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelFromClass{TClass}"/> class.
		/// </summary>
		/// <param name="path">The path.</param>
		/// <param name="worksheet">The worksheet.</param>
		public ReadExcelFromClass(string path, string worksheet)
			: base(path, worksheet, true) { }

		/// <summary>
		/// Initializes a new instance of the <see cref="ReadExcelFromClass{TClass}"/> class.
		/// </summary>
		/// <param name="stream">The stream.</param>
		/// <param name="worksheet">The worksheet.</param>
		public ReadExcelFromClass(Stream stream, string worksheet)
			: base(stream, worksheet, true) { }

		/// <summary>
		/// Reads the Excel.
		/// </summary>
		/// <returns></returns>
		/// <exception cref="System.ArgumentNullException">Columns are required to map them from the excel sheet!</exception>
		public override TClass[] Read()
		{
			if (Columns == null || Columns.Any() == false)
			{
				throw new ArgumentNullException("Columns are required to map them from the excel sheet!");
			}

			DataTable data = base.GetDataTableFromExcel();

			List<TClass> instances = new List<TClass>();
			Type type = typeof(TClass);
			PropertyInfo[] propertyInfos = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

			int index;
			string value;
			object instance;
			PropertyInfo propertyInfo;
			TypeConverter converter;
			foreach (DataRow row in data.Rows)
			{
				instance = Activator.CreateInstance(type);

				foreach (KeyValuePair<string[], Expression<Func<TClass, dynamic>>> field in this.Columns)
				{
					// Find index of the column
					foreach (string columnName in field.Key)
					{
						index = data.Columns.IndexOf(columnName);

						if (index >= 0)
						{
							//TODO: Check if string.IsNullOrWhiteSpace() for column?

							// Then try to get the field
							value = row.Field<string>(columnName);

							if (string.IsNullOrWhiteSpace(value)) continue;

							//TODO: Helper class for this? Because this is used a lot...
							//TODO: type.GetProperty(field.Value.GetArgumentName()) instead?
							propertyInfo = propertyInfos.Single(o => o.Name.Equals(field.Value.GetArgumentName(), StringComparison.OrdinalIgnoreCase));

							// Convert the type
							if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
							{
								// Check if it is a swiss date
								if (DateTimeParser.IsSwissDate(value))
								{
									propertyInfo.SetValue(instance, DateTimeParser.ParseSwissDateTimeString(value), null);
									break; // stop searching other column name
								}
							}

							converter = TypeDescriptor.GetConverter(propertyInfo.PropertyType);

							propertyInfo.SetValue(instance, converter.ConvertFromInvariantString(value), null);

							break; // stop searching other column name
						}
					}
				}

				instances.Add(instance as TClass);
			}

			return instances.ToArray();
		}

	}
}