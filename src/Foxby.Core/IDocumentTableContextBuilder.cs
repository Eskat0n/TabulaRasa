using System;

namespace Foxby.Core
{
	public interface IDocumentTableContextBuilder
	{
		IDocumentTagContextBuilder Table(Action<IDocumentTableSchemeBuilder> options, Action<IDocumentTableRowsBuilder> rows);
		IDocumentTagContextBuilder BorderNone { get; }
	}
}