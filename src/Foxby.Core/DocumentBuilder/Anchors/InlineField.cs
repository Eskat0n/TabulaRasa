namespace TabulaRasa.DocumentBuilder.Anchors
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Extensions;

    internal class InlineField : SdtField<SdtRun, SdtContentRun>
	{
	    private readonly RunProperties _runProperties;

	    private InlineField(string elementName, SdtRun sdtElement)
			: base(elementName, sdtElement)
		{
			var runProperties = sdtElement
				.Descendants<RunProperties>()
				.FirstOrDefault();

	        _runProperties = runProperties != null ? runProperties.CloneElement() : null;
		}

	    internal RunProperties RunProperties
        {
            get
            {
                return _runProperties == null
                        ? new RunProperties()
                        : _runProperties.CloneElement();
            }
        }

		internal static IEnumerable<InlineField> Get(WordprocessingDocument document, string fieldName)
		{
			return Get(document, fieldName, (fn, run) => new InlineField(fn, run));
		}
	}
}