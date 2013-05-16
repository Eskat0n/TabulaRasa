namespace TabulaRasa.Tests.DocumentBuilder
{
    using System.Linq;
    using NUnit.Framework;
    using Properties;
    using TabulaRasa.DocumentBuilder.Anchors;
    using MetaObjects;                                    

    [TestFixture]
    public class DocumentPlaceholderTests
	{
		[Test]
		public void NewPlaceholderNameShouldBeCorrect()
		{
			var placeholder = new DocumentPlaceholder("TEST_PLACEHOLDER");

			Assert.AreEqual("TEST_PLACEHOLDER", placeholder.Name);
		}

		[Test]
		public void NewPlaceholderOpeningEnclosureShouldBeCorrect()
		{
			var placeholder = new DocumentPlaceholder("TEST_PLACEHOLDER");

			Assert.AreEqual("{{TEST_PLACEHOLDER}}", placeholder.OpeningName);
		}

		[Test]
		public void NewPlaceholderClosingEnclosureShouldBeCorrect()
		{
			var placeholder = new DocumentPlaceholder("TEST_PLACEHOLDER");

			Assert.AreEqual("{{/TEST_PLACEHOLDER}}", placeholder.ClosingName);
		}

		[Test]
		public void NewPlaceholderCreationCreatesItsParagraphs()
		{
			var placeholder = new DocumentPlaceholder("TEST_PH");

			Assert.NotNull(placeholder.Opening);
			Assert.Null(placeholder.Opening.Parent);
			Assert.NotNull(placeholder.Closing);
			Assert.Null(placeholder.Closing.Parent);
		}

		[Test]
		public void ManyPlaceholdersGettingFromDocumentCorrect()
		{
			using (var document = new DocxDocument(Resources.WithManyPlaceholders))
			{
				var placeholders = DocumentPlaceholder.Get(document.GetWordDocument(), "INNER");

				Assert.NotNull(placeholders);
				Assert.AreEqual(3, placeholders.Count());
			}
		}

		[Test]
		public void PlaceholderGettingWhichNotExistsReturnsEmpty()
		{
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var placeholders = DocumentPlaceholder.Get(document.GetWordDocument(), "NON_EXISTING");

				Assert.IsEmpty(placeholders);
			}
		}
	}
}