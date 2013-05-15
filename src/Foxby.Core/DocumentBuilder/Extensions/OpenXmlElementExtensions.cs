namespace TabulaRasa.DocumentBuilder.Extensions
{
    using DocumentFormat.OpenXml;

    internal static class OpenXmlElementExtensions
	{
		public static TElement CloneElement<TElement>(this TElement element)
			where TElement : OpenXmlElement
		{
			return (TElement) element.Clone();
		}
	}
}