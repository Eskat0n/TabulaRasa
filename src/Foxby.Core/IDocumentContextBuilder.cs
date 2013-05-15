namespace TabulaRasa
{
    using System;

    /// <summary>
	/// Provide base methods for changing content of block anchors and OpenXML elements
	/// </summary>
	public interface IDocumentContextBuilder : IDocumentContextFormattableBuilder
	{
		/// <summary>
		/// Explicitly indicates begin of editable area in document
		/// </summary>
		IDocumentContextBuilder EditableStart();

		/// <summary>
		/// Explicitly indicates end of editable area in document
		/// </summary>
		IDocumentContextBuilder EditableEnd();

		/// <summary>
		/// Inserts placeholder with specified <paramref name="placeholderName"/> and content declared by delegate <paramref name="options"/>
		/// </summary>
		/// <param name="placeholderName">Placeholder name</param>
		/// <param name="options">Delegate which contains code filling placeholder content</param>
		IDocumentContextBuilder Placeholder(string placeholderName, Action<IDocumentContextBuilder> options = null);

		/// <summary>
		/// Inserts placeholder with specified <paramref name="placeholderName"/> and text content from <paramref name="contentLines"/>
		/// </summary>
		/// <param name="placeholderName">Placeholder name</param>
		/// <param name="contentLines">Text content</param>
		IDocumentContextBuilder Placeholder(string placeholderName, params string[] contentLines);

		/// <summary>
		/// Appends text from <paramref name="contentLines"/>
		/// </summary>
		/// <param name="contentLines">Text content</param>
		IDocumentContextBuilder AddText(params string[] contentLines);

		/// <summary>
		/// Inserts image specified by <paramref name="content"/>
		/// </summary>
		/// <param name="content">Image binary content</param>
		/// <param name="contentType">MIME type</param>
        IDocumentContextBuilder Image(byte[] content, string contentType);
	}
}