using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Extensions;
using Xunit;

namespace Foxby.Core.Tests.DocumentBuilder.Extensions
{
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