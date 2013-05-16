namespace TabulaRasa.Tests.DocumentBuilder
{
    using System.Linq;
    using NUnit.Framework;
    using Properties;
    using TabulaRasa.DocumentBuilder.Anchors;
    using MetaObjects;

    [TestFixture]
    public class DocumentTagTests
	{
		[Test]
		public void NewTagNameShouldBeCorrect()
		{
			var tag = new DocumentTag("TEST_TAG");

			Assert.AreEqual("TEST_TAG", tag.Name);
		}

		[Test]
		public void NewTagOpeningTagNameShoulbBeEnclosedInWavedBrackets()
		{
			var tag = new DocumentTag("TEST_TAG");

            Assert.AreEqual("{TEST_TAG}", tag.OpeningName);
		}

		[Test]
		public void NewTagClosingTagNameShouldBeEnclosedInWavedBracketsWithSlash()
		{
			var tag = new DocumentTag("TEST_TAG");

            Assert.AreEqual("{/TEST_TAG}", tag.ClosingName);
		}

		[Test]
		public void NewTagCreationCreatesItsParagraphs()
		{
			var tag = new DocumentTag("TEST_TAG");

			Assert.NotNull(tag.Opening);
			Assert.Null(tag.Opening.Parent);
			Assert.NotNull(tag.Closing);
			Assert.Null(tag.Closing.Parent);
		}

		[Test]
		public void ManyTagsGettingFromDocumentCorrect()
		{
			using (var document = new DocxDocument(Resources.WithManyTags))
			{
				var tags = DocumentTag.Get(document.GetWordDocument(), "SUB");

				Assert.NotNull(tags);
                Assert.AreEqual(3, tags.Count());
			}
		}

		[Test]
		public void TagGettingWhichNotExistsReturnsEmpty()
		{
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var tags = DocumentTag.Get(document.GetWordDocument(), "NON_EXISTING");

                Assert.IsEmpty(tags);
			}
		}
	}
}