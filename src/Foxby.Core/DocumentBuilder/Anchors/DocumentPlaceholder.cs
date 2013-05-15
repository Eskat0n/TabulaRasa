namespace TabulaRasa.DocumentBuilder.Anchors
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    ///<summary>
	/// Represents inline anchor (placeholder)
	///</summary>
	internal class DocumentPlaceholder : AnchorElement<Run>
	{
		///<summary>
		/// ctor
		///</summary>
		///<param name="placeholderName">Name of new placeholder</param>
		public DocumentPlaceholder(string placeholderName) 
			: base(placeholderName, "{{{{{0}}}}}", "{{{{/{0}}}}}")
		{
			Opening = new Run(new Text(OpeningName)) {RunProperties = new RunProperties(new Vanish())};			
			Closing = new Run(new Text(ClosingName)) {RunProperties = new RunProperties(new Vanish())};
		}

		private DocumentPlaceholder(Run opening, Run closing, string placeholderName)
			: base(placeholderName, "{{{{{0}}}}}", "{{{{/{0}}}}}")
		{
			Opening = opening;
			Closing = closing;
		}

		internal static IEnumerable<DocumentPlaceholder> Get(WordprocessingDocument document, string placeholderName)
		{
			var placeholder = new DocumentPlaceholder(placeholderName);

			var result = new List<DocumentPlaceholder>();

			var openings = GetElements(document).Where(x => x.InnerText == placeholder.OpeningName);
			var closings = GetElements(document).Where(x => x.InnerText == placeholder.ClosingName);

			for (var i = 0; i < openings.Count(); i++)
				if (i < closings.Count())
				{
					var opening = openings.ElementAt(i);
					var closing = closings.ElementAt(i);

					if (opening.RunProperties != null)
						opening.RunProperties.Vanish = new Vanish();
					else
						opening.RunProperties = new RunProperties(new Vanish());
					if (closing.RunProperties != null)
						closing.RunProperties.Vanish = new Vanish();
					else
						closing.RunProperties = new RunProperties(new Vanish());

					result.Add(new DocumentPlaceholder(opening, closing, placeholderName));
				}

			return result;
		}
	}
}