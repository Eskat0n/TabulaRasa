using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Extensions;

namespace Foxby.Core.DocumentBuilder.Anchors
{
	internal class BlockField : SdtField<SdtBlock, SdtContentBlock>
	{
		private readonly ParagraphStyleId _paragraphStyleId;

		private BlockField(string elementName, SdtBlock sdtElement)
			: base(elementName, sdtElement)
		{
			var paragraphProperties = sdtElement
				.Descendants<ParagraphProperties>()
				.FirstOrDefault();

			_paragraphStyleId = paragraphProperties != null && paragraphProperties.ParagraphStyleId != null
			                    	? paragraphProperties.ParagraphStyleId.CloneElement()
			                    	: null;
		}

		internal ParagraphStyleId ParagraphStyleId
		{
			get
			{
				return _paragraphStyleId == null
				       	? null
				       	: _paragraphStyleId.CloneElement();
			}
		}

		internal static IEnumerable<BlockField> Get(WordprocessingDocument document, string fieldName)
		{
			return Get(document, fieldName, (fn, run) => new BlockField(fn, run));
		}
	}
}