using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Foxby.Core.DocumentBuilder.Anchors
{
	public abstract class AnchorElement<TElement>
		where TElement : OpenXmlElement
	{
		public string Name { get; private set; }
		public string OpeningName { get; private set; }
		public string ClosingName { get; private set; }
		public TElement Opening { get; protected set; }
		public TElement Closing { get; protected set; }

		protected AnchorElement(string elementName, string openingNameFormat, string closingNameFormat)
		{
			Name = elementName;
			OpeningName = string.Format(openingNameFormat, elementName);
			ClosingName = string.Format(closingNameFormat, elementName);
		}

		protected static IEnumerable<TElement> GetElements(WordprocessingDocument document)
		{
			return document.MainDocumentPart.Document
				.Descendants()
				.OfType<TElement>();
		}

		public void Remove()
		{
			Opening.Remove();
			Closing.Remove();
		}
	}
}