namespace Foxby.Core.MetaObjects.Extensions
{
	public static class StringFormatExtensions
	{
		public static Format Bold(this string @this)
		{
			Format format = @this;
			return format.Bold();
		}

		public static Format Underlined(this string @this)
		{
			Format format = @this;
			return format.Underlined();
		}

		public static Format Italic(this string @this)
		{
			Format format = @this;
			return format.Italic();
		}
	}
}