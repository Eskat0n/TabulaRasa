namespace TabulaRasa.DocumentBuilder
{
    using System;
    using System.Linq;
    using Anchors;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Format = MetaObjects.Format;

    internal abstract class DocxDocumentBlockContextBuilderBase : DocxDocumentBuilderBase, IDocumentTagContextBuilder
	{
		private OpenXmlElement[] _contextParagraphPrependedElements;
		private ParagraphProperties _contextParagraphProperties;
		private TableProperties _contextTableProperties;
		private bool _isEditable;

		protected DocxDocumentBlockContextBuilderBase(WordprocessingDocument document)
			: base(document)
		{
		}

		public IDocumentTagContextBuilder BorderNone
		{
			get
			{
				_contextTableProperties = new TableProperties(new Border {Val = BorderValues.None});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Left
		{
			get
			{
				_contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Left});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Center
		{
			get
			{
				_contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Center});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Right
		{
			get
			{
				_contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Right});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Both
		{
			get
			{
				_contextParagraphProperties = new ParagraphProperties(new Justification {Val = JustificationValues.Both});
				return this;
			}
		}

		public IDocumentParagraphFormattableBuilder Indent
		{
			get
			{
				_contextParagraphPrependedElements = new OpenXmlElement[] {new Run(new TabChar())};
				return this;
			}
		}

		public IDocumentTagContextBuilder EditableStart()
		{
			if (_isEditable)
				return this;

			AppendElements(CreatePermStart());
			SaveDocument();

			_isEditable = true;

			return this;
		}

		public IDocumentTagContextBuilder EditableEnd()
		{
			if (_isEditable == false)
				return this;

			AppendElements(CreatePermEnd());
			SaveDocument();

			_isEditable = false;

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
			var builder = new DocxDocumentParagraphContextBuilder(Document, _contextParagraphProperties, _contextParagraphPrependedElements);
			
			options(builder);

			AppendElements(builder.ToElement());

			SaveDocument();

			_contextParagraphProperties = null;
			_contextParagraphPrependedElements = null;

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
			var tableContextBuilder = new DocxDocumentTableSchemeBuilder(Document, _contextTableProperties);

			header.Invoke(tableContextBuilder);
			rows.Invoke(tableContextBuilder);

			AppendElements(tableContextBuilder.AggregatedContent.ToArray());

			_contextTableProperties = null;

			return this;
		}

		protected abstract void AppendElements(params OpenXmlElement[] openXmlElements);
	}
}