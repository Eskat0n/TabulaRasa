using System;

namespace Foxby.Core
{
	/// <summary>
	/// Provide methods for top level operations
	/// </summary>
	public interface IDocumentBuilder
	{
		/// <summary>
		/// Sets internal content to tag with name <paramref name="tagName"/>
		/// </summary>
		/// <param name="tagName">Tag name from template</param>
		/// <param name="options">Delegate which contains code filling tag content</param>
		IDocumentBuilder Tag(string tagName, Action<IDocumentTagContextBuilder> options);

		/// <summary>
		/// Sets internal content to placeholder with name <paramref name="placeholderName"/>
		/// </summary>
		/// <param name="placeholderName">Placeholder name from template</param>
		/// <param name="options">Delegate which contains code filling placeholder content</param>
		/// <param name="preservePlaceholder">Indicates whether to remove placeholder after setting its content. Default is true.</param>
		IDocumentBuilder Placeholder(string placeholderName, Action<IDocumentContextBuilder> options, bool preservePlaceholder = true);

		/// <summary>
		/// Sets internal content to block field (sdt element) with name <paramref name="fieldName"/>
		/// </summary>
		/// <param name="fieldName">Field name from template</param>
		/// <param name="options">Delegate which contains code filling field content</param>
		IDocumentBuilder BlockField(string fieldName, Action<IDocumentTagContextBuilder> options);

		/// <summary>
		/// Hide or display content of tag with name <paramref name="tagName"/>
		/// </summary>
		/// <param name="tagName">Tag name from template</param>
		/// <param name="visible">Hide content if false; otherwise, show content</param>
        void SetVisibilityTag(string tagName, bool visible);

		/// <summary>
		/// Validate OpenXML document against a schema
		/// </summary>
		/// <returns>True if content is valid; otherwise, false</returns>
	    bool Validate();
		
		/// <summary>
		/// Serialize OpenXML document as binary array
		/// </summary>
		byte[] ToArray();
	}
}