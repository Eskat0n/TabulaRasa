namespace TabulaRasa.Tests.DocumentBuilder
{
    using System.Linq;
    using TabulaRasa.Tests.Properties;
    using Xunit;
    using TabulaRasa.DocumentBuilder.Anchors;
    using TabulaRasa.MetaObjects;

    public class DocumentPlaceholderTests
	{
		[Fact]
		public void NewPlaceholderNameShouldBeCorrect()
		{
			var placeholder = new DocumentPlaceholder("TEST_PLACEHOLDER");

			Assert.Equal("TEST_PLACEHOLDER", placeholder.Name);
		}

		[Fact]
		public void NewPlaceholderOpeningEnclosureShouldBeCorrect()
		{
			var placeholder = new DocumentPlaceholder("TEST_PLACEHOLDER");

			Assert.Equal("{{TEST_PLACEHOLDER}}", placeholder.OpeningName);
		}

		[Fact]
		public void NewPlaceholderClosingEnclosureShouldBeCorrect()
		{
			var placeholder = new DocumentPlaceholder("TEST_PLACEHOLDER");

			Assert.Equal("{{/TEST_PLACEHOLDER}}", placeholder.ClosingName);
		}

		[Fact]
		public void NewPlaceholderCreationCreatesItsParagraphs()
		{
			var placeholder = new DocumentPlaceholder("TEST_PH");

			Assert.NotNull(placeholder.Opening);
			Assert.Null(placeholder.Opening.Parent);
			Assert.NotNull(placeholder.Closing);
			Assert.Null(placeholder.Closing.Parent);
		}

		[Fact]
		public void ManyPlaceholdersGettingFromDocumentCorrect()
		{
			using (var document = new DocxDocument(Resources.WithManyPlaceholders))
			{
				var placeholders = DocumentPlaceholder.Get(document.GetWordDocument(), "INNER");

				Assert.NotNull(placeholders);
				Assert.Equal(3, placeholders.Count());
			}
		}

		[Fact]
		public void PlaceholderGettingWhichNotExistsReturnsEmpty()
		{
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var placeholders = DocumentPlaceholder.Get(document.GetWordDocument(), "NON_EXISTING");

				Assert.Empty(placeholders);
			}
		}
	}
}