using System;
using Foxby.Core.MetaObjects;

namespace Foxby.Core
{
	public interface IDocumentParagraphFormattableBuilder : IDocumentFormattableBuilder<IDocumentParagraphFormattableBuilder>
	{
		IDocumentTagContextBuilder Paragraph(params string[] contentLines);
		IDocumentTagContextBuilder Paragraph(Format content);
		IDocumentTagContextBuilder Paragraph(Action<IDocumentContextBuilder> options);

		IDocumentParagraphFormattableBuilder Indent { get; }
	}
}