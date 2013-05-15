namespace TabulaRasa
{
    using System;

    ///<summary>
	/// Provides methods for inserting tables
	///</summary>
	public interface IDocumentTableContextBuilder
	{
		///<summary>
		/// Inserts new table with specified <paramref name="header"/> and <paramref name="rows"/>
		///</summary>
		///<param name="header">Delegate which contains code defining table header</param>
		///<param name="rows">Delegate which contains code filling table rows content</param>
		IDocumentTagContextBuilder Table(Action<IDocumentTableSchemeBuilder> header, Action<IDocumentTableRowsBuilder> rows);

		///<summary>
		/// Indicates whether preceding table will have no borders
		///</summary>
		IDocumentTagContextBuilder BorderNone { get; }
	}
}