namespace TabulaRasa
{
    using System;

    ///<summary>
	/// Provide base methods for changing content of tags
	///</summary>
	public interface IDocumentTagContextBuilder : IDocumentParagraphFormattableBuilder, IDocumentTableContextBuilder
	{
		///<summary>
		/// Explicitly indicates begin of editable area in document
		///</summary>
		IDocumentTagContextBuilder EditableStart();

		/// <summary>
		/// Explicitly indicates end of editable area in document
		/// </summary>
		IDocumentTagContextBuilder EditableEnd();

		///<summary>
		/// Appends empty line to tag
		///</summary>
		IDocumentTagContextBuilder EmptyLine();

		///<summary>
		/// Appends specified <paramref name="count"/> of empty lines to tag
		///</summary>
		///<param name="count"></param>
		IDocumentTagContextBuilder EmptyLine(int count);

		///<summary>
		/// Appends new children tag with specified <paramref name="tagName"/> and content
		///</summary>
		///<param name="tagName">New tag name</param>
		///<param name="options">Delegate which contains code filling new tag content</param>
		IDocumentTagContextBuilder AppendTag(string tagName, Action<IDocumentTagContextBuilder> options);

		///<summary>
		/// Appends ordered list to tag
		///</summary>
		///<param name="options">Delegate which contains code filling list with items</param>
		IDocumentTagContextBuilder OrderedList(Action<IDocumentOrderedListBuilder> options);
	}
}