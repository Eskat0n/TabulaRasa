namespace TabulaRasa.DocumentBuilder
{
    using System.Collections.Generic;
    using Anchors;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using Extensions;

    internal class DocxDocumentTagContextBuilder : DocxDocumentBlockContextBuilderBase
	{
		private readonly IEnumerable<DocumentTag> _documentTags;

		public DocxDocumentTagContextBuilder(WordprocessingDocument document, string tagName)
			: base(document)
		{
			_documentTags = DocumentTag.Get(document, tagName);

			foreach (DocumentTag documentTag in _documentTags)
				ClearBetweenElements(documentTag.Opening, documentTag.Closing);
			SaveDocument();
		}

		protected override void AppendElements(params OpenXmlElement[] openXmlElements)
		{
			foreach (var documentTag in _documentTags)
				foreach (var openXmlElement in openXmlElements)
					documentTag.Closing.InsertBeforeSelf(openXmlElement.CloneElement());
		}
	}
}