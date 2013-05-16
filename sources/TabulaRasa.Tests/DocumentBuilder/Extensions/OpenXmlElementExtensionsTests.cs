namespace TabulaRasa.Tests.DocumentBuilder.Extensions
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using NUnit.Framework;
    using TabulaRasa.DocumentBuilder.Extensions;

    [TestFixture]
    public class OpenXmlElementExtensionsTests
	{
		[Test]
		public void CloneElementCorrect()
		{
			var initial = new Run();
			var cloned = initial.CloneElement();

			Assert.False(ReferenceEquals(initial, cloned));
		}
	}
}