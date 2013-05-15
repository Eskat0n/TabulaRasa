using Foxby.Core.DocumentBuilder;
using Foxby.Core.MetaObjects;
using Foxby.Core.Tests.EqualityComparers;
using TabulaRasa.Tests.Properties;
using Xunit;

namespace Foxby.Core.Tests.DocumentBuilder
{
	public class DocxDocumentBuilderFieldTests
	{
		[Fact]
		public void CanSetContentToBlockFieldConsistingOfTwoParagraphs()
		{
			using (var expected = new DocxDocument(Resources.WithTwoParagraphsInBlockField))
			using (var document = new DocxDocument(Resources.WithSdtElements))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.BlockField("BlockField", x => x.Paragraph("Первый").Paragraph("Второй"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}
		
		[Fact]
		public void CanSetContentToInlineFieldConsistingOfTwoRuns()
		{
			using (var expected = new DocxDocument(Resources.WithTwoRunsInInlineField))
			using (var document = new DocxDocument(Resources.WithSdtElements))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.InlineField("InlineField", x => x.Text("Первый").Text("Второй"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void PropertiesAppliedToBlockFieldShouldPersistAfterSettingContent()
		{
			using (var expected = new DocxDocument(Resources.WithStyledSdtElementsContentInserted))
			using (var document = new DocxDocument(Resources.WithStyledSdtElements))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.BlockField("BlockField", x => x.Paragraph("Первый").Paragraph("Второй"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}			
		}
		
		[Fact]
		public void PropertiesAppliedToInlineFieldShouldPersistAfterSettingContent()
		{
			using (var expected = new DocxDocument(Resources.WithStyledInlineSdtElementContentInserted))
			using (var document = new DocxDocument(Resources.WithStyledInlineSdtElement))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.InlineField("InlineField", x => x.Text("Первый").Text("Второй"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}			
		}
		
        [Fact]
		public void PropertiesAppliedToInlineFieldInHeadersAndFooters()
        {
            using (var expected = new DocxDocument(Resources.FieldsInHeadersAndFootersReplaced))
            using (var document = new DocxDocument(Resources.FieldsInHeadersAndFooters))
			{
				var builder = new DocxDocumentBuilder(document);

                builder.InlineField("Signer.ShortNameThisOrSubstitute", x => x.Text("Первый").Text("Второй"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}			
		}
	}
}