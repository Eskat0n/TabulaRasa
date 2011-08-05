using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Foxby.Core.DocumentBuilder
{
	///<summary>
	/// Base class for anchors content builders with embedded buffer
	///</summary>
	public abstract class DocxDocumentAggregationBuilder : DocxDocumentBuilderBase
	{
		internal readonly List<OpenXmlElement> Aggregation = new List<OpenXmlElement>();

		protected DocxDocumentAggregationBuilder(WordprocessingDocument document)
			: base(document)
		{
		}

		///<summary>
		/// Gets aggregated content of anchor have been built by this builder
		///</summary>
		public IEnumerable<OpenXmlElement> AggregatedContent
		{
			get { return Aggregation; }
		}
	}
}