namespace TabulaRasa.DocumentBuilder
{
    using System.Collections.Generic;

    /// <summary>
	/// List of tags for show and hide 
	/// </summary>
	public class TagVisibilityOptions
	{
		/// <summary>
		/// Tag names for hide
		/// </summary>
		public IEnumerable<string> HiddenTagNames { get; private set; }

		/// <summary>
		/// Tag name for show
		/// </summary>
		public string VisibleTagName { get; private set; }
		
		/// <summary>
		/// ctor
		/// </summary>
		/// <param name="visibleTagName">Tag name for show</param>
		/// <param name="hiddenTagNames">Tag names for hide</param>
		public TagVisibilityOptions(string visibleTagName, IEnumerable<string> hiddenTagNames)
		{
			VisibleTagName = visibleTagName;
			HiddenTagNames = hiddenTagNames;
		}
	}
}