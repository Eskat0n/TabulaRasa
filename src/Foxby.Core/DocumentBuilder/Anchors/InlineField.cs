using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder.Anchors
{
	internal class InlineField : SdtField<SdtRun, SdtContentRun>
	{
		private InlineField(string elementName, SdtRun sdtElement)
			: base(elementName, sdtElement)
		{
		}

		internal static IEnumerable<InlineField> Get(WordprocessingDocument document, string fieldName)
		{
			return Get(document, fieldName, (fn, run) => new InlineField(fn, run));
		}
	}
}