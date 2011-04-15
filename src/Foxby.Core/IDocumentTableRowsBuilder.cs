using System;

namespace Foxby.Core
{
	public interface IDocumentTableRowsBuilder
	{
		IDocumentTableRowsBuilder Row(params string[] rowContent);
		IDocumentTableRowsBuilder Row(params Action<ICellContextBuilder>[] optionsParams);
	}
}