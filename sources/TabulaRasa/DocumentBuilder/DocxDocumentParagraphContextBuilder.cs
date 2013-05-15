namespace TabulaRasa.DocumentBuilder
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal class DocxDocumentParagraphContextBuilder : DocxDocumentContextBuilderBase
	{
		private readonly IEnumerable<OpenXmlElement> _prependedElements;
		private readonly ParagraphProperties _properties;

		public DocxDocumentParagraphContextBuilder(WordprocessingDocument document, ParagraphProperties properties, IEnumerable<OpenXmlElement> prependedElements)
			: base(document)
		{
			_prependedElements = prependedElements;
			_properties = properties;
		}

		protected override RunProperties RunProperties
		{
			get { return new RunProperties(new Vanish()); }
		}

		public OpenXmlElement ToElement()
		{
			var paragraphContent = new List<OpenXmlElement>();
			if (_prependedElements != null)
				paragraphContent.AddRange(_prependedElements);

			paragraphContent.AddRange(AggregatedContent);

			return new Paragraph(paragraphContent) {ParagraphProperties = _properties};
		}
	}
}