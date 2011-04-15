namespace Foxby.Core
{
	public interface IDocumentFormattableBuilder<out TBuilder>
	{
		TBuilder Left { get; }
		TBuilder Center { get; }
		TBuilder Right { get; }
		TBuilder Both { get; }
	}
}