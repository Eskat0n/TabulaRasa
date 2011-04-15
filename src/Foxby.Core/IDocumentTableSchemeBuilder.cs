namespace Foxby.Core
{
	public interface IDocumentTableSchemeBuilder
	{
		IDocumentTableSchemeBuilder Column(string columnName);
		IDocumentTableSchemeBuilder Column(string columnName, float widthInPercents);
	}
}