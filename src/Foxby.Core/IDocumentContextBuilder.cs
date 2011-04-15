using System;

namespace Foxby.Core
{
	public interface IDocumentContextBuilder : IDocumentContextFormattableBuilder
	{
		IDocumentContextBuilder EditableStart();

		IDocumentContextBuilder EditableEnd();

		IDocumentContextBuilder Placeholder(string placeholderName, Action<IDocumentContextBuilder> options = null);

		IDocumentContextBuilder Placeholder(string placeholderName, params string[] contentLines);

		IDocumentContextBuilder AddText(params string[] contentLines);
	}
}