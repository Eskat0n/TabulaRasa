using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder.Anchors
{
	internal abstract class SdtField<TSdtElement, TSdtContentElement> : AnchorElement<TSdtElement>
		where TSdtElement : SdtElement
		where TSdtContentElement : OpenXmlElement
	{
		protected SdtField(string elementName, TSdtElement sdtElement)
			: base(elementName, null, null)
		{
			Opening = sdtElement;
			Closing = sdtElement;
		}

		internal TSdtContentElement ContentWrapper
		{
			get { return Opening.GetFirstChild<TSdtContentElement>(); }
		}
		
		internal IEnumerable<OpenXmlElement> Content
		{
			get { return Opening.Elements<TSdtContentElement>().SelectMany(x => x.ChildElements); }
		}

		protected static IEnumerable<TSdtField> Get<TSdtField>(WordprocessingDocument document, string fieldName, Func<string, TSdtElement, TSdtField> factory)
			where TSdtField : SdtField<TSdtElement, TSdtContentElement>
		{
			return GetElements(document)
				.Where(x => x.SdtProperties.GetFirstChild<SdtAlias>().Val.Value == fieldName)
				.Select(x => factory(fieldName, x))
				.ToList();
		}
	}
}