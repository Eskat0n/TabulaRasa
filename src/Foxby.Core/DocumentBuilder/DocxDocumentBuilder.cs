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
	public sealed class DocxDocumentBuilder : DocxDocumentBuilderBase, IDocumentBuilder
	{
		private readonly DocxDocument docxDocument;

		private DocxDocumentBuilder(DocxDocument docxDocument)
			: base(docxDocument.GetWordDocument())
		{
			this.docxDocument = docxDocument;

			MergeVanishedRuns();
		}

		public static IDocumentBuilder Create(DocxDocument docxDocument)
		{
			return new DocxDocumentBuilder(docxDocument);
		}

		public IDocumentBuilder Tag(string tagName, Action<IDocumentTagContextBuilder> options)
		{
			options(new DocxDocumentTagContextBuilder(Document, tagName));
			return this;
		}

		public IDocumentBuilder Placeholder(string placeholderName, Action<IDocumentContextBuilder> options, bool isUpdatable = true)
		{
			var documentPlaceholders = DocumentPlaceholder.Get(Document, placeholderName);

			foreach (var documentPlaceholder in documentPlaceholders)
				ClearBetweenElements(documentPlaceholder.Opening, documentPlaceholder.Closing);
			SaveDocument();

			var placeholderContextBuilder = new DocxDocumentPlaceholderContextBuilder(Document, documentPlaceholders.Count() == 0
			                                                                                    	? new RunProperties()
			                                                                                    	: documentPlaceholders.First().Opening.RunProperties);
			options(placeholderContextBuilder);

			foreach (var documentPlaceholder in documentPlaceholders)
			{
				foreach (var contentElement in placeholderContextBuilder.AggregatedContent)
					documentPlaceholder.Closing.InsertBeforeSelf(contentElement.CloneElement());
				if (isUpdatable == false)
					documentPlaceholder.Remove();
			}
			SaveDocument();

			return this;
		}

	    public bool Validate()
	    {
	        var errorInfos = new OpenXmlValidator(FileFormatVersions.Office2007).Validate(Document);
	        return errorInfos.Count() == 0;
	    }

	    public byte[] ToArray()
		{
			return docxDocument.ToArray();
		}

		private void MergeVanishedRuns()
		{
			var vanishedRuns = Document.MainDocumentPart.Document.Descendants().OfType<Run>()				
				.Where(SplittedRunPredicate).ToList();

			var vanishedNeighbours = new List<List<Run>>();
			foreach (var vanishedRun in vanishedRuns)
				if (vanishedNeighbours.Count == 0 || vanishedNeighbours.Last().Last().NextSibling<Run>() != vanishedRun)
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