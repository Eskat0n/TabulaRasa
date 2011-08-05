using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder.Anchors
{
	///<summary>
	/// Represents block anchor (tag)
	///</summary>
	internal class DocumentTag : AnchorElement<Paragraph>
	{
		///<summary>
		/// ctor
		///</summary>
		///<param name="tagName">Name of new tag</param>
		public DocumentTag(string tagName)
			: base(tagName, "{{{0}}}", "{{/{0}}}")
		{
			var openingParagraphContent = new Run(new RunProperties(new Vanish()), new Text(OpeningName));
			Opening = new Paragraph(openingParagraphContent) {ParagraphProperties = new ParagraphProperties(new RunProperties(new Vanish()))};
			var closingParagraphContent = new Run(new RunProperties(new Vanish()), new Text(ClosingName));
			Closing = new Paragraph(closingParagraphContent) {ParagraphProperties = new ParagraphProperties(new RunProperties(new Vanish()))};
		}

		private DocumentTag(Paragraph opening, Paragraph closing, string tagName)
			: base(tagName, "{{{0}}}", "{{/{0}}}")
		{
			Opening = opening;
			Closing = closing;
		}

		internal static IEnumerable<DocumentTag> Get(WordprocessingDocument document, string tagName)
		{
			var tag = new DocumentTag(tagName);

			var result = new List<DocumentTag>();

			var openings = GetElements(document).Where(x => x.InnerText == tag.OpeningName);
			var closings = GetElements(document).Where(x => x.InnerText == tag.ClosingName);

			for (var i = 0; i < openings.Count(); i++)
				if (i < closings.Count())
					result.Add(new DocumentTag(openings.ElementAt(i), closings.ElementAt(i), tagName));

			return result;
		}
	}
}