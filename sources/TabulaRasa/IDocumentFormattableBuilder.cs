namespace TabulaRasa
{
	/// <summary>
	/// Provides alignment options for block OpenXML elements
	/// </summary>
	/// <typeparam name="TBuilder">Concrete builder type</typeparam>
	public interface IDocumentFormattableBuilder<out TBuilder>
	{
		/// <summary>
		/// Aligns element to left
		/// </summary>
		TBuilder Left { get; }

		/// <summary>
		/// Aligns element to center
		/// </summary>
		TBuilder Center { get; }

		/// <summary>
		/// Aligns element to right
		/// </summary>
		TBuilder Right { get; }

		/// <summary>
		/// Aligns element to justify
		/// </summary>
		TBuilder Both { get; }
	}
}