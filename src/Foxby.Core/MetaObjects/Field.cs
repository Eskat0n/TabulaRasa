namespace TabulaRasa.MetaObjects
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

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
			get { return Get<SdtAlias>().Val.Value; }
			set { Get<SdtAlias>().Val = new StringValue(value); }
		}

		/// <summary>
		/// Name of tag of underlying <see cref="SdtElement"/>
		/// </summary>
		public string Tag
		{
			get { return Get<Tag>().Val.Value; }
			set { Get<Tag>().Val = new StringValue(value); }
		}

		private T Get<T>() where T : StringType, new()
		{
			SdtProperties properties = _element.SdtProperties;

			return properties.GetFirstChild<T>() ??
			       properties.AppendChild(new T {Val = new StringValue()});
		}
	}
}