using System;

namespace Foxby.Core
{
	///<summary>
	/// Provides methods for inserting table rows with content
	///</summary>
	public interface IDocumentTableRowsBuilder
	{
		///<summary>
		/// Appends new row with specified cell <paramref name="content"/>
		///</summary>
		///<param name="content">Text content where each string represents single cell content</param>
		IDocumentTableRowsBuilder Row(params string[] content);

		///<summary>
		/// Appends new row with content built by delegates <paramref name="options"/>
		///</summary>
		///<param name="options">Delegates which contains code filling table row content</param>
		IDocumentTableRowsBuilder Row(params Action<ICellContextBuilder>[] options);
	}
}