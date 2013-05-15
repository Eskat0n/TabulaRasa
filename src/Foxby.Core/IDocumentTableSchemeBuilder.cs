namespace TabulaRasa
{
	///<summary>
	/// Provides methods for defining table header and columns metadata
	///</summary>
	public interface IDocumentTableSchemeBuilder
	{
		///<summary>
		/// Defines column with specified <paramref name="columnName"/>
		///</summary>
		///<param name="columnName">Column name as displayed in header</param>
		IDocumentTableSchemeBuilder Column(string columnName);

		///<summary>
		/// Defines column with specified <paramref name="columnName"/> and <paramref name="widthInPercents"/>
		///</summary>
		///<param name="columnName">Column name as displayed in header</param>
		///<param name="widthInPercents">Width for column in percents (must be set for all columns)</param>
		IDocumentTableSchemeBuilder Column(string columnName, float widthInPercents);
	}
}