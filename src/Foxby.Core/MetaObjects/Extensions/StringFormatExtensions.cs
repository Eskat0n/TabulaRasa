namespace TabulaRasa.MetaObjects.Extensions
{
	///<summary>
	/// Provides extension methods for specifying formatting options for text strings
	///</summary>
	public static class StringFormatExtensions
	{
		///<summary>
		/// Applies bold format to string
		///</summary>
		///<param name="this"></param>
		public static Format Bold(this string @this)
		{
			Format format = @this;
			return format.Bold();
		}

		/// <summary>
		/// Applies underlined format to string
		/// </summary>
		/// <param name="this"></param>
		public static Format Underlined(this string @this)
		{
			Format format = @this;
			return format.Underlined();
		}

		/// <summary>
		/// Applies italic format to string
		/// </summary>
		/// <param name="this"></param>
		public static Format Italic(this string @this)
		{
			Format format = @this;
			return format.Italic();
		}
	}
}