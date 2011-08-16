using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.MetaObjects
{
	/// <summary>
	/// Class which holds basic information about <see cref="SdtElement"/>
	/// </summary>
	public class Field
	{
		private readonly SdtElement _element;

		internal Field(SdtElement element)
		{
			_element = element;
		}

		/// <summary>
		/// Name of underlying <see cref="SdtElement"/>
		/// </summary>
		public string Name
		{
			get { return _element.SdtProperties.GetFirstChild<SdtAlias>().Val.Value; }
			set { _element.SdtProperties.GetFirstChild<SdtAlias>().Val = new StringValue(value); }
		}

		/// <summary>
		/// Name of tag of underlying <see cref="SdtElement"/>
		/// </summary>
		public string Tag
		{
			get { return _element.SdtProperties.GetFirstChild<Tag>().Val.Value; }
			set { _element.SdtProperties.GetFirstChild<Tag>().Val = new StringValue(value); }
		}
	}
}