using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder
{
	internal class DocxDocumentCellContextBuilder : DocxDocumentContextBuilderBase, ICellContextBuilder
	{
		private readonly TableCellProperties contextTableCellProperties;
		private ParagraphProperties paragraphProperties;

		public DocxDocumentCellContextBuilder(WordprocessingDocument document)
			: base(document)
		{
			contextTableCellProperties = new TableCellProperties(new TableCellWidth {Type = TableWidthUnitValues.Auto});
		}

		public DocxDocumentCellContextBuilder(WordprocessingDocument document, TableCellProperties cellProperties)
            : base(document)
        {
            contextTableCellProperties = cellProperties;
        }

		public ICellContextBuilder Left
		{
			get
			{
				paragraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Left});
				return this;
			}
		}

		public ICellContextBuilder Center
		{
			get
			{
				paragraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Center });
				return this;
			}
		}

		public ICellContextBuilder Right
		{
			get
			{
				paragraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Right});
				return this;
			}
		}

		public ICellContextBuilder Both
		{
			get
			{
				paragraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Both });
				return this;
			}
		}

		public void Cell(string content)
		{
			Aggregation.Add(new Run(new Text(content) {Space = SpaceProcessingModeValues.Preserve}));
		}

        public void Cell(Action<IDocumentContextBuilder> content)
        {
            var documentContextBuilder = new DocxDocumentParagraphContextBuilder(Document, null, null);

            content(documentContextBuilder);

            Aggregation.AddRange(documentContextBuilder.Aggregation);
        }

	    public OpenXmlElement ToElement()
		{
			var tableCell = new TableCell
			                	{
			                		TableCellProperties = contextTableCellProperties
			                	};

			tableCell.AppendChild(new Paragraph(Aggregation.ToArray())
			                      	{
			                      		ParagraphProperties = paragraphProperties
			                      	});

			return tableCell;
		}

		protected override RunProperties RunProperties
		{
			get { return new RunProperties(new Vanish()); }
		}
	}
}