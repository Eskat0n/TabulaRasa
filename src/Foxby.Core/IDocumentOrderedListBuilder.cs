namespace TabulaRasa
{
    using System;

    /// <summary>
	/// Provider methods for creating ordered list items
	/// </summary>
	public interface IDocumentOrderedListBuilder
	{
		/// <summary>
		/// Create one ordered list item
		/// </summary>
		/// <param name="contentLines">Text lines</param>
		IDocumentOrderedListBuilder Item(params string[] contentLines);

		/// <summary>
		/// Create one ordered list item
		/// </summary>
		/// <param name="options">Delegate which contains code filling list item content</param>
		IDocumentOrderedListBuilder Item(Action<IDocumentContextBuilder> options);
	}
}