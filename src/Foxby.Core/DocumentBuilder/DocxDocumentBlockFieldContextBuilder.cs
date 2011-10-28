using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Anchors;
using Foxby.Core.DocumentBuilder.Extensions;

namespace Foxby.Core.DocumentBuilder
{
	internal class DocxDocumentBlockFieldContextBuilder : DocxDocumentBlockContextBuilderBase
	{
		private readonly IEnumerable<BlockField> _documentFields;

		public DocxDocumentBlockFieldContextBuilder(WordprocessingDocument document, string fieldName)
			: base(document)
		{
			_documentFields = BlockField.Get(document, fieldName);

			foreach (var blockField in _documentFields)
				blockField.ContentWrapper.RemoveAllChildren();
			SaveDocument();
		}

		protected override void AppendElements(params OpenXmlElement[] openXmlElements)
		{
			foreach (var blockField in _documentFields)
			{
				var cloned = openXmlElements
					.Select(x => x.CloneElement())
					.ToList();

				var paragraphs = cloned
					.OfType<Paragraph>();

				foreach (var paragraph in paragraphs)
					if (paragraph.ParagraphProperties == null)
						paragraph.ParagraphProperties = new ParagraphProperties(blockField.ParagraphStyleId);
					else
						paragraph.ParagraphProperties.ParagraphStyleId = blockField.ParagraphStyleId;

				blockField.ContentWrapper.Append(cloned);
			}
		}
	}
}