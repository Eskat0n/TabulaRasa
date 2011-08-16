using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.MetaObjects
{
	public class Field
	{
		private readonly SdtElement _element;

		internal Field(SdtElement element)
		{
			_element = element;
		}

		public string Name
		{
			get { return _element.SdtProperties.GetFirstChild<SdtAlias>().Val.Value; }
			set { _element.SdtProperties.GetFirstChild<SdtAlias>().Val = new StringValue(value); }
		}

		public string Tag
		{
			get { return _element.SdtProperties.GetFirstChild<Tag>().Val.Value; }
			set { _element.SdtProperties.GetFirstChild<Tag>().Val = new StringValue(value); }
		}
	}
}