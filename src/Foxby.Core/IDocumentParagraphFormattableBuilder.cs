using System;
using Foxby.Core.MetaObjects;

namespace Foxby.Core
{
	///<summary>
	/// Provides methods for inserting paragraphs with specified indentation and format
	///</summary>
	public interface IDocumentParagraphFormattableBuilder : IDocumentFormattableBuilder<IDocumentParagraphFormattableBuilder>
	{
		///<summary>
		/// Appends paragraph with specified text <paramref name="content"/> to current tag
		///</summary>
		///<param name="content">Text content of new paragraph</param>
		IDocumentTagContextBuilder Paragraph(params string[] content);

		///<summary>
		/// Appends paragraph with specified formatted text <paramref name="content"/> to current tag
		///</summary>
		///<param name="content">Formatted text content of new paragraph</param>
		IDocumentTagContextBuilder Paragraph(Format content);

		///<summary>
		/// Appends paragraph with specified by builder <paramref name="options"/> to current tag
		///</summary>
		///<param name="options">Delegate which contains code filling paragraph content</param>
		IDocumentTagContextBuilder Paragraph(Action<IDocumentContextBuilder> options);

		///<summary>
		/// Specifies one tab-width indentation preceding the paragraph beign inserted
		///</summary>
		IDocumentParagraphFormattableBuilder Indent { get; }
	}
}