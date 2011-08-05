using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder
{
	internal class DocxDocumentParagraphContextBuilder : DocxDocumentContextBuilderBase
	{
		private readonly IEnumerable<OpenXmlElement> prependedElements;
		private readonly ParagraphProperties properties;

		public DocxDocumentParagraphContextBuilder(WordprocessingDocument document, ParagraphProperties properties, IEnumerable<OpenXmlElement> prependedElements)
			: base(document)
		{
			this.prependedElements = prependedElements;
			this.properties = properties;
		}

		protected override RunProperties RunProperties
		{
			get { return new RunProperties(new Vanish()); }
		}

		public OpenXmlElement ToElement()
		{
			var paragraphContent = new List<OpenXmlElement>();
			if (prependedElements != null)
				paragraphContent.AddRange(prependedElements);

			paragraphContent.AddRange(AggregatedContent);

			return new Paragraph(paragraphContent) {ParagraphProperties = properties};
		}
	}
}