using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder
{
	public class DocxDocumentOrderedListBuilder : DocxDocumentAggregationBuilder, IDocumentOrderedListBuilder
	{
		private readonly int numberingId;

		public DocxDocumentOrderedListBuilder(WordprocessingDocument document)
			: base(document)
		{
			var numberingProperties = Document.MainDocumentPart.Document.
				Descendants()
				.OfType<NumberingProperties>()
				.Select(x => x.NumberingId.Val);

			numberingId = numberingProperties.Any() == false
			              	? 1
			              	: (numberingProperties.Max(x => x.Value) + 1);
		}

		public IDocumentOrderedListBuilder Item(params string[] contentLines)
		{
			return Item(x => x.AddText(contentLines));
		}

		public IDocumentOrderedListBuilder Item(Action<IDocumentContextBuilder> options)
		{
			var builder = new DocxDocumentParagraphContextBuilder(Document, CreateParagraphProperties(), null);
			
			options(builder);

			var paragraph = builder.ToElement();

			Aggregation.Add(paragraph);

			return this;
		}

		private ParagraphProperties CreateParagraphProperties()
		{
			return new ParagraphProperties(new NumberingProperties(new NumberingLevelReference {Val = 0},
			                                                       new NumberingId {Val = numberingId}),
			                               new Indentation {Left = "720", Hanging = "0"});
		}
	}
}