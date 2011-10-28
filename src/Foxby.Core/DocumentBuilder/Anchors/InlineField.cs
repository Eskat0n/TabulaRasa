using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Extensions;

namespace Foxby.Core.DocumentBuilder.Anchors
{
	internal class InlineField : SdtField<SdtRun, SdtContentRun>
	{
		private readonly RunStyle _runStyle;

		private InlineField(string elementName, SdtRun sdtElement)
			: base(elementName, sdtElement)
		{
			var runProperties = sdtElement
				.Descendants<RunProperties>()
				.FirstOrDefault();

			_runStyle = runProperties != null
			            	? runProperties.RunStyle.CloneElement()
			            	: null;
		}

		internal RunStyle RunStyle
		{
			get { return _runStyle.CloneElement(); }
		}

		internal static IEnumerable<InlineField> Get(WordprocessingDocument document, string fieldName)
		{
			return Get(document, fieldName, (fn, run) => new InlineField(fn, run));
		}
	}
}