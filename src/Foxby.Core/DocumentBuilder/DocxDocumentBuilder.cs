using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Anchors;
using Foxby.Core.DocumentBuilder.Extensions;
using Foxby.Core.MetaObjects;

namespace Foxby.Core.DocumentBuilder
{
	/// <summary>
	/// Operates with content of tags and placeholders in OpenXML template
	/// </summary>
	public class DocxDocumentBuilder : DocxDocumentBuilderBase, IDocumentBuilder
	{
		private readonly DocxDocument docxDocument;
        private readonly TagVisibilityOptions tagVisibilityOptions;

		/// <summary>
		/// Create builder for template
		/// </summary>
		/// <param name="docxDocument">Template document</param>
		/// <param name="tagVisibilityOptions">List of tags for show and hide</param>
	    public DocxDocumentBuilder(DocxDocument docxDocument, TagVisibilityOptions tagVisibilityOptions = null)
			: base(docxDocument.GetWordDocument())
		{
			this.docxDocument = docxDocument;
            this.tagVisibilityOptions = tagVisibilityOptions;

			MergeVanishedRuns();
		}

		/// <summary>
		/// Factory method
		/// </summary>
		/// <param name="docxDocument">Template document</param>
        public static IDocumentBuilder Create(DocxDocument docxDocument)
		{
			return new DocxDocumentBuilder(docxDocument);
		}

		public IDocumentBuilder Tag(string tagName, Action<IDocumentTagContextBuilder> options)
		{
			options(new DocxDocumentTagContextBuilder(Document, tagName));
			return this;
		}

		public IDocumentBuilder Placeholder(string placeholderName, Action<IDocumentContextBuilder> options, bool preservePlaceholder = true)
		{
			var documentPlaceholders = DocumentPlaceholder.Get(Document, placeholderName);

			foreach (var documentPlaceholder in documentPlaceholders)
				ClearBetweenElements(documentPlaceholder.Opening, documentPlaceholder.Closing);

			SaveDocument();

			foreach (var documentPlaceholder in documentPlaceholders)
			{
				var builder = new DocxDocumentPlaceholderContextBuilder(Document, documentPlaceholders.Any() == false
				                                                                  	? new RunProperties()
				                                                                  	: documentPlaceholder.Opening.RunProperties);
                options(builder);
                foreach (var contentElement in builder.AggregatedContent)
					documentPlaceholder.Closing.InsertBeforeSelf(contentElement.CloneElement());
				if (preservePlaceholder == false)
					documentPlaceholder.Remove();
			}

			SaveDocument();

			return this;
		}

		public IDocumentBuilder BlockField(string fieldName, Action<IDocumentTagContextBuilder> options)
		{
			options(new DocxDocumentBlockFieldContextBuilder(Document, fieldName));
			return this;
		}

		public IDocumentBuilder InlineField(string fieldName, Action<IDocumentContextBuilder> options)
		{
			var inlineFields = Anchors.InlineField.Get(Document, fieldName);

			foreach (var inlineField in inlineFields)
				inlineField.ContentWrapper.RemoveAllChildren();

			SaveDocument();

			foreach (var inlineField in inlineFields)
			{
				var builder = new DocxDocumentPlaceholderContextBuilder(Document, new RunProperties());

				options(builder);

				var cloned = builder.AggregatedContent.Select(x => x.CloneElement());
				inlineField.ContentWrapper.Append(cloned);
			}

			SaveDocument();

			return this;
		}

		public void SetVisibilityTag(string tagName, bool visible)
        {
            docxDocument.SetTagVisibility(tagName, visible);
        }

	    public bool Validate()
	    {
	        var errorInfos = new OpenXmlValidator(FileFormatVersions.Office2007).Validate(Document);
	        return !errorInfos.Any();
	    }

	    public byte[] ToArray()
		{
            if (tagVisibilityOptions != null)
                docxDocument.SetTagVisibility(tagVisibilityOptions);
			return docxDocument.ToArray();
		}

		private void MergeVanishedRuns()
		{
			var vanishedRuns = Document.MainDocumentPart.Document.Descendants().OfType<Run>()				
				.Where(SplittedRunPredicate).ToList();

			var vanishedNeighbours = new List<List<Run>>();
			foreach (var vanishedRun in vanishedRuns)
				if (!vanishedNeighbours.Any() || vanishedNeighbours.Last().Last().NextSibling<Run>() != vanishedRun)
					vanishedNeighbours.Add(new List<Run> {vanishedRun});
				else
					vanishedNeighbours.Last().Add(vanishedRun);

			vanishedNeighbours.Where(x => x.Count > 1)
				.ToList()
				.ForEach(x =>
				         	{
				         		x.First()
				         			.OfType<Text>()
				         			.Last().Text = x.Aggregate(string.Empty, (a, c) => a + c.InnerText);
				         		x.RemoveAt(0);
				         		x.ForEach(z => z.Remove());
				         	});
		}

		private static bool SplittedRunPredicate(Run run)
		{
			var placeholderRegex = new Regex(@"{{/?.+}}");
			return run.RunProperties != null && run.RunProperties.Vanish != null && placeholderRegex.IsMatch(run.InnerText) == false;
		}
	}
}