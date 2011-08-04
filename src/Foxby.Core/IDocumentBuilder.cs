using System;

namespace Foxby.Core
{
	public interface IDocumentBuilder
	{
		IDocumentBuilder Tag(string tagName, Action<IDocumentTagContextBuilder> options);

		IDocumentBuilder Placeholder(string placeholderName, Action<IDocumentContextBuilder> options, bool isUpdatable = true);

        void SetVisibilityTag(string tagName, bool visible);

	    bool Validate();
		
		byte[] ToArray();
	}
}