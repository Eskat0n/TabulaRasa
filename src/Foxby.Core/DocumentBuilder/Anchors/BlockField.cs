namespace TabulaRasa.DocumentBuilder.Anchors
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Extensions;

    internal class BlockField : SdtField<SdtBlock, SdtContentBlock>
	{
		private readonly ParagraphProperties _paragraphProperties;

		private BlockField(string elementName, SdtBlock sdtElement)
			: base(elementName, sdtElement)
		{
			var paragraphProperties = sdtElement
				.Descendants<ParagraphProperties>()
				.FirstOrDefault();

		    _paragraphProperties = paragraphProperties != null ? paragraphProperties.CloneElement() : null;
		}

	    internal ParagraphProperties ParagraphProperties
	    {
	        get { return _paragraphProperties == null ? new ParagraphProperties() : _paragraphProperties.CloneElement(); }
	    }

		internal static IEnumerable<BlockField> Get(WordprocessingDocument document, string fieldName)
		{
			return Get(document, fieldName, (fn, run) => new BlockField(fn, run));
		}
	}
}