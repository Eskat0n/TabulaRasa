namespace TabulaRasa.Tests.DocumentBuilder.Extensions
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using Xunit;
    using TabulaRasa.DocumentBuilder.Extensions;

    public class OpenXmlElementExtensionsTests
	{
		[Fact]
		public void CloneElementCorrect()
		{
			var initial = new Run();
			var cloned = initial.CloneElement();

			Assert.False(ReferenceEquals(initial, cloned));
		}
	}
}