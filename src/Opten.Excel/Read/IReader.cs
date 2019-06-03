namespace Opten.Excel.Read
{
	/// <summary>
	/// The Reader.
	/// </summary>
	/// <typeparam name="TOutput">The type of the output.</typeparam>
	public interface IReader<TOutput> where TOutput : class
	{

		/// <summary>
		/// Reads the Excel or CSV.
		/// </summary>
		/// <returns></returns>
		TOutput[] Read();

	}
}