using System.Linq;
using Foxby.Core.DocumentBuilder.Anchors;
using Foxby.Core.MetaObjects;
using TabulaRasa.Tests.Properties;
using Xunit;

namespace Foxby.Core.Tests.DocumentBuilder
{
	public class DocumentTagTests
	{
		[Fact]
		public void NewTagNameShouldBeCorrect()
		{
			var tag = new DocumentTag("TEST_TAG");

			Assert.Equal("TEST_TAG", tag.Name);
		}

		[Fact]
		public void NewTagOpeningTagNameShoulbBeEnclosedInWavedBrackets()
		{
			var tag = new DocumentTag("TEST_TAG");

			Assert.Equal("{TEST_TAG}", tag.OpeningName);
		}

		[Fact]
		public void NewTagClosingTagNameShouldBeEnclosedInWavedBracketsWithSlash()
		{
			var tag = new DocumentTag("TEST_TAG");

			Assert.Equal("{/TEST_TAG}", tag.ClosingName);
		}

		[Fact]
		public void NewTagCreationCreatesItsParagraphs()
		{
			var tag = new DocumentTag("TEST_TAG");

			Assert.NotNull(tag.Opening);
			Assert.Null(tag.Opening.Parent);
			Assert.NotNull(tag.Closing);
			Assert.Null(tag.Closing.Parent);
		}

		[Fact]
		public void ManyTagsGettingFromDocumentCorrect()
		{
			using (var document = new DocxDocument(Resources.WithManyTags))
			{
				var tags = DocumentTag.Get(document.GetWordDocument(), "SUB");

				Assert.NotNull(tags);
				Assert.Equal(3, tags.Count());
			}
		}

		[Fact]
		public void TagGettingWhichNotExistsReturnsEmpty()
		{
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var tags = DocumentTag.Get(document.GetWordDocument(), "NON_EXISTING");

				Assert.Empty(tags);
			}
		}
	}
}