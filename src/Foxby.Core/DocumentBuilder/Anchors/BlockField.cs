using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder.Anchors
{
	internal class BlockField : SdtField<SdtBlock, SdtContentBlock>
	{
		private BlockField(string elementName, SdtBlock sdtElement)
			: base(elementName, sdtElement)
		{
		}

		internal static IEnumerable<BlockField> Get(WordprocessingDocument document, string fieldName)
		{
			return Get(document, fieldName, (fn, run) => new BlockField(fn, run));
		}
	}
}