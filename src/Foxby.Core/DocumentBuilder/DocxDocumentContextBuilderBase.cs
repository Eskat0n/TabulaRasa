using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Anchors;
using Foxby.Core.DocumentBuilder.Extensions;

namespace Foxby.Core.DocumentBuilder
{
	public abstract class DocxDocumentContextBuilderBase : DocxDocumentAggregationBuilder, IDocumentContextBuilder
	{
		private readonly ICollection<OpenXmlElement> runPropertiesElements = new List<OpenXmlElement>();

		private bool isEditable;

		protected DocxDocumentContextBuilderBase(WordprocessingDocument document)
			: base(document)
		{
		}

		protected abstract RunProperties RunProperties { get; }

		public IDocumentContextBuilder EditableStart()
		{
			if (isEditable)
				return this;

			Aggregation.Add(CreatePermStart());
			isEditable = true;
			return this;
		}

		public IDocumentContextBuilder EditableEnd()
		{
			if (isEditable == false)
				return this;

			Aggregation.Add(CreatePermEnd());
			isEditable = false;
			return this;
		}

		public IDocumentContextBuilder Line(string contentLine)
		{
			return Text(contentLine + Environment.NewLine);
		}

		public IDocumentContextFormattableBuilder Bold
		{
			get
			{
				runPropertiesElements.Add(new Bold());
				return this;
			}
		}

		public IDocumentContextFormattableBuilder Italic
		{
			get
			{
				runPropertiesElements.Add(new Italic());
				return this;
			}
		}

		public IDocumentContextFormattableBuilder Underlined
		{
			get
			{
				runPropertiesElements.Add(new Underline {Val = UnderlineValues.Single});
				return this;
			}
		}

		public virtual IDocumentContextBuilder Placeholder(string placeholderName, Action<IDocumentContextBuilder> options)
		{
			var newPlaceholder = new DocumentPlaceholder(placeholderName);

			Aggregation.Add(newPlaceholder.Opening);

			if (options != null)
			{
				var placeholderContextBuilder = new DocxDocumentPlaceholderContextBuilder(Document, RunProperties);
				options(placeholderContextBuilder);
				Aggregation.AddRange(placeholderContextBuilder.AggregatedContent);
			}

			Aggregation.Add(newPlaceholder.Closing);

			return this;
		}

		public IDocumentContextBuilder Placeholder(string placeholderName, params string[] contentLines)
		{
			return Placeholder(placeholderName, x => x.Text(contentLines));
		}

		public IDocumentContextBuilder AddText(params string[] contentLines)
		{
			Aggregation.AddRange(CreateTextContent(contentLines));

			return this;
		}

		public IDocumentContextBuilder Text(params string[] contentLines)
		{
			RunProperties properties = RunProperties.CloneElement();
			properties.Vanish = null;
			properties.Append(runPropertiesElements);
			Aggregation.AddRange(CreateTextContent(contentLines, properties));
			runPropertiesElements.Clear();

			return this;
			}
		}
}