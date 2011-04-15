namespace Foxby.Core
{
	public interface IDocumentContextFormattableBuilder
	{
		IDocumentContextBuilder Text(params string[] contentLines);
		IDocumentContextBuilder Line(string contentLine);

		IDocumentContextFormattableBuilder Bold { get; }
		IDocumentContextFormattableBuilder Italic { get; }
		IDocumentContextFormattableBuilder Underlined { get; }
	}
}