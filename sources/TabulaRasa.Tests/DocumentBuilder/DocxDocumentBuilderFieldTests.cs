namespace TabulaRasa.Tests.DocumentBuilder
{
    using EqualityComparers;
    using NUnit.Framework;
    using Properties;
    using TabulaRasa.DocumentBuilder;
    using MetaObjects;

    [TestFixture]
    public class DocxDocumentBuilderFieldTests
	{
		[Test]
		public void CanSetContentToBlockFieldConsistingOfTwoParagraphs()
		{
			using (var expected = new DocxDocument(Resources.WithTwoParagraphsInBlockField))
			using (var document = new DocxDocument(Resources.WithSdtElements))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.BlockField("BlockField", x => x.Paragraph("Первый").Paragraph("Второй"));

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expected, document));
			}
		}
		
		[Test]
		public void CanSetContentToInlineFieldConsistingOfTwoRuns()
		{
			using (var expected = new DocxDocument(Resources.WithTwoRunsInInlineField))
			using (var document = new DocxDocument(Resources.WithSdtElements))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.InlineField("InlineField", x => x.Text("Первый").Text("Второй"));

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expected, document));
			}
		}

		[Test]
		public void PropertiesAppliedToBlockFieldShouldPersistAfterSettingContent()
		{
			using (var expected = new DocxDocument(Resources.WithStyledSdtElementsContentInserted))
			using (var document = new DocxDocument(Resources.WithStyledSdtElements))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.BlockField("BlockField", x => x.Paragraph("Первый").Paragraph("Второй"));

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expected, document));
			}			
		}
		
		[Test]
		public void PropertiesAppliedToInlineFieldShouldPersistAfterSettingContent()
		{
			using (var expected = new DocxDocument(Resources.WithStyledInlineSdtElementContentInserted))
			using (var document = new DocxDocument(Resources.WithStyledInlineSdtElement))
			{
				var builder = new DocxDocumentBuilder(document);

				builder.InlineField("InlineField", x => x.Text("Первый").Text("Второй"));

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expected, document));
			}			
		}
		
        [Test]
		public void PropertiesAppliedToInlineFieldInHeadersAndFooters()
        {
            using (var expected = new DocxDocument(Resources.FieldsInHeadersAndFootersReplaced))
            using (var document = new DocxDocument(Resources.FieldsInHeadersAndFooters))
			{
				var builder = new DocxDocumentBuilder(document);

                builder.InlineField("Signer.ShortNameThisOrSubstitute", x => x.Text("Первый").Text("Второй"));

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expected, document));
			}			
		}
	}
}