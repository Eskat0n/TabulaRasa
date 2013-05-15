namespace TabulaRasa
{
	/// <summary>
	/// Provide methods for edit or format text content of inline elements
	/// </summary>
	public interface IDocumentContextFormattableBuilder
	{
		/// <summary>
		/// Sets text lines
		/// </summary>
		/// <param name="contentLines">Text lines</param>
		IDocumentContextBuilder Text(params string[] contentLines);

		/// <summary>
		/// Sets text line
		/// </summary>
		/// <param name="contentLine">Text line</param>
		IDocumentContextBuilder Line(string contentLine);

		/// <summary>
		/// Format text as bold
		/// </summary>
		IDocumentContextFormattableBuilder Bold { get; }

		/// <summary>
		/// Format text as italic
		/// </summary>
		IDocumentContextFormattableBuilder Italic { get; }

		/// <summary>
		/// Format text as underlined
		/// </summary>
		IDocumentContextFormattableBuilder Underlined { get; }
	}
}