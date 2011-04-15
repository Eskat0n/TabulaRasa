using System;

namespace Foxby.Core
{
	public interface IDocumentTagContextBuilder : IDocumentParagraphFormattableBuilder, IDocumentTableContextBuilder
	{
		IDocumentTagContextBuilder EditableStart();
		IDocumentTagContextBuilder EditableEnd();

		IDocumentTagContextBuilder EmptyLine();
		IDocumentTagContextBuilder EmptyLine(int count);

		IDocumentTagContextBuilder NewTag(string tagName, Action<IDocumentTagContextBuilder> options);
		IDocumentTagContextBuilder OrderedList(Action<IDocumentOrderedListBuilder> options);
	}
}