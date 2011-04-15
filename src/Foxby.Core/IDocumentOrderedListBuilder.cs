using System;

namespace Foxby.Core
{
	public interface IDocumentOrderedListBuilder
	{
		IDocumentOrderedListBuilder Item(params string[] contentLines);
		IDocumentOrderedListBuilder Item(Action<IDocumentContextBuilder> options);
	}
}