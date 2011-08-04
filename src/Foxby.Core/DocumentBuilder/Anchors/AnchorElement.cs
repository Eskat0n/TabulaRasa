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
		    var ofType = document.MainDocumentPart.Document
		        .Descendants()
		        .OfType<TElement>().ToList();


		    var elementsInHeaders = new List<TElement>();
		    foreach (var headerPart in document.MainDocumentPart.HeaderParts)
		        elementsInHeaders.AddRange(headerPart.RootElement.Descendants().OfType<TElement>());

		    ofType.AddRange(elementsInHeaders);

            var elementsInFooters = new List<TElement>();
            foreach (var footerPart in document.MainDocumentPart.FooterParts)
                elementsInFooters.AddRange(footerPart.RootElement.Descendants().OfType<TElement>());

            ofType.AddRange(elementsInFooters);

            return ofType;
		}

	    public void Remove()
		{
			Opening.Remove();
			Closing.Remove();
		}
	}
}