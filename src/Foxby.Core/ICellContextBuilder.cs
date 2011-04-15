using System;

namespace Foxby.Core
{
	public interface ICellContextBuilder : IDocumentContextBuilder, IDocumentFormattableBuilder<ICellContextBuilder>
	{
		void Cell(string content);
	    void Cell(Action<IDocumentContextBuilder> content);
	}
}