namespace TabulaRasa.DocumentBuilder
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;

    internal abstract class DocxDocumentAggregationBuilder : DocxDocumentBuilderBase
	{
		internal readonly List<OpenXmlElement> Aggregation = new List<OpenXmlElement>();

		protected DocxDocumentAggregationBuilder(WordprocessingDocument document)
			: base(document)
		{
		}

		public IEnumerable<OpenXmlElement> AggregatedContent
		{
			get { return Aggregation; }
		}
	}
}