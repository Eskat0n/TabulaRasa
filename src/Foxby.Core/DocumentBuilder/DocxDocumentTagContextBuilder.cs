using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Anchors;
using Foxby.Core.DocumentBuilder.Extensions;
using Format = Foxby.Core.MetaObjects.Format;

namespace Foxby.Core.DocumentBuilder
{
	internal class DocxDocumentTagContextBuilder : DocxDocumentBuilderBase, IDocumentTagContextBuilder
	{
		private readonly IEnumerable<DocumentTag> documentTags;

		private OpenXmlElement[] contextParagraphPrependedElements;
		private ParagraphProperties contextParagraphProperties;
		private TableProperties contextTableProperties;

		private bool isEditable;

		public DocxDocumentTagContextBuilder(WordprocessingDocument document, string tagName)
			: base(document)
		{
			documentTags = DocumentTag.Get(document, tagName);

			foreach (DocumentTag documentTag in documentTags)
				ClearBetweenElements(documentTag.Opening, documentTag.Closing);
			SaveDocument();
		}

		public IDocumentTagContextBuilder EditableStart()
		{
			if (isEditable)
				return this;

			AppendElements(CreatePermStart());
			SaveDocument();

			isEditable = true;

			return this;
		}

		public IDocumentTagContextBuilder EditableEnd()
		{
			if (isEditable == false)
				return this;

			AppendElements(CreatePermEnd());
			SaveDocument();

			isEditable = false;

			return this;
		}

		public IDocumentTagContextBuilder EmptyLine()
		{
			return EmptyLine(1);
		}

		public IDocumentTagContextBuilder EmptyLine(int count)
		{
			var builder = new DocxDocumentParagraphContextBuilder(Document, null, null);
			
			for (int i = 0; i < count; i++)
				AppendElements(builder.ToElement());

			SaveDocument();

			return this;
		}

		public IDocumentTagContextBuilder Paragraph(params string[] content)
		{
			return Paragraph(x => x.AddText(content));
		}

        public IDocumentTagContextBuilder Paragraph(Format content)
		{
			return Paragraph(content.Invoke);
		}

		public IDocumentTagContextBuilder Paragraph(Action<IDocumentContextBuilder> options)
		{
			var builder = new DocxDocumentParagraphContextBuilder(Document, contextParagraphProperties, contextParagraphPrependedElements);
			
			options(builder);

			AppendElements(builder.ToElement());

			SaveDocument();

			contextParagraphProperties = null;
			contextParagraphPrependedElements = null;

			return this;
		}

		public IDocumentTagContextBuilder OrderedList(Action<IDocumentOrderedListBuilder> options)
		{
			var builder = new DocxDocumentOrderedListBuilder(Document);
			
			options(builder);

			AppendElements(builder.AggregatedContent.ToArray());

			SaveDocument();

			return this;
		}

		public IDocumentTagContextBuilder AppendTag(string tagName, Action<IDocumentTagContextBuilder> options)
		{
			var tag = new DocumentTag(tagName);

			AppendElements(tag.Opening, tag.Closing);
			
			SaveDocument();

			options(new DocxDocumentTagContextBuilder(Document, tagName));
			return this;
		}

		public IDocumentTagContextBuilder Table(Action<IDocumentTableSchemeBuilder> header, Action<IDocumentTableRowsBuilder> rows)
		{
			var tableContextBuilder = new DocxDocumentTableSchemeBuilder(Document, contextTableProperties);

			header.Invoke(tableContextBuilder);
			rows.Invoke(tableContextBuilder);

			AppendElements(tableContextBuilder.AggregatedContent.ToArray());

			contextTableProperties = null;

			return this;
		}

		public IDocumentTagContextBuilder BorderNone
		{
			get
			{
				contextTableProperties = new TableProperties(new Border {Val = BorderValues.None});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Left
		{
			get
			{
				contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Left});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Center
		{
			get
			{
				contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Center});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Right
		{
			get
			{
				contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Right});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Both
		{
			get
			{
				contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Both});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Indent
		{
			get
			{
				contextParagraphPrependedElements = new OpenXmlElement[] {new Run(new TabChar())};
				return this;
			}
		}

		private void AppendElements(params OpenXmlElement[] openXmlElements)
		{
			foreach (DocumentTag documentTag in documentTags)
				foreach (OpenXmlElement openXmlElement in openXmlElements)
					documentTag.Closing.InsertBeforeSelf(openXmlElement.CloneElement());
		}
	}
}