using DocumentFormat.OpenXml;

namespace Foxby.Core.DocumentBuilder.Extensions
{
	public static class OpenXmlElementExtensions
	{
		public static TElement CloneElement<TElement>(this TElement element)
			where TElement : OpenXmlElement
		{
			return (TElement) element.Clone();
		}
	}
}