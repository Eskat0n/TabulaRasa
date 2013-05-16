namespace TabulaRasa.Tests
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;
    using EqualityComparers;
    using NUnit.Framework;
    using Properties;
    using System;
    using MetaObjects;

    [TestFixture]
    public class DocxDocumentTests
    {
        [Test]
        public void DocumentsWithDifferentInnerXmlMustBeNotEqual()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            using (var otherDocument = new DocxDocument(Resources.DocumentWithParagraph))
            {
                var comparer = new DocxDocumentEqualityComparer();

                Assert.Throws<AssertionException>(() => comparer.Equals(document, otherDocument));
            }
        }

        [Test]
        public void DocumentsWithEqualInnerXmlMustBeEqual()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            using (var otherDocument = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(document, otherDocument));
            }
        }

        [Test]
        [Ignore("Несовпадающие rsidR у элементов параграфов документа")]
        public void TestCleanContent()
        {
            using (var withParagraph = new DocxDocument(Resources.DocumentWithParagraph))
            using (var withOutParagraph = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                withParagraph.CleanContent("Edit");

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(withOutParagraph, withParagraph));
            }
        }

        [Test]
        public void ReplacePairsFromDictionary()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithTitle))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithReplacedTitle))
            {
                initialDocument.Replace(new Dictionary<string, string>
				                        	{
				                        		{"Title", "Hello"}
				                        	});
                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expectedDocument, initialDocument));
            }
        }

        [Test]
        public void ReplaceSingleTagWithTextBlocksDoesNothing()
        {
            using (var initialDocument = new DocxDocument(Resources.WithMainContentSingleTag))
            {
                var content = new[] { new TextBlock("Контент документа") };
                initialDocument.Replace("MAIN_CONTENT", content);

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(initialDocument, initialDocument));
            }
        }

        [Test]
        public void ReplaceOpenCloseTagWithTextBlocksEditable()
        {
            using (var initialDocument = new DocxDocument(Resources.WithMainContentTag))
            using (var expectedDocument = new DocxDocument(Resources.WithMainContentInserted))
            {
                var content = new[] { new TextBlock("Контент документа") };
                initialDocument.Replace("MAIN_CONTENT", content);

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expectedDocument, initialDocument));
            }
        }

        [Test]
        public void InsertTagContentToOpenCloseTag()
        {
            using (var initialDocument = new DocxDocument(Resources.WithMainContentTag))
            using (var expectedDocument = new DocxDocument(Resources.WithMainContentInserted))
            {
                var content = new[] { new TextBlock("Контент документа") };
                initialDocument.InsertTagContent("MAIN_CONTENT", content);

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expectedDocument, initialDocument));
            }
        }

        [Test]
        public void UnprotectRemovesWrightProtectionFromFile()
        {
            using (var @protected = new DocxDocument(Resources.Protected))
            using (var @unprotected = new DocxDocument(Resources.Unprotected))
            {
                @protected.Unprotect();

                var unprotectedOuterXml = @unprotected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml;
                var protectedOuterXml = @protected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml;

                Assert.IsTrue(unprotectedOuterXml.Equals(protectedOuterXml, StringComparison.InvariantCultureIgnoreCase));
            }
        }

        [Test]
        public void ProtectAddWrightProtectionToFile()
        {
            using (var @protected = new DocxDocument(Resources.Protected))
            using (var @unprotected = new DocxDocument(Resources.Unprotected))
            {
                @unprotected.Protect();

                Assert.AreEqual(@protected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml, @unprotected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml);
            }
        }

        [Test]
        public void SetCustomPropertyToDocument()
        {
            using (var withoutAttributes = new DocxDocument(Resources.DocumentWithoutAttributes))
            using (var withAttribute = new DocxDocument(Resources.DocumentWithAttribute))
            {
                withoutAttributes.SetCustomProperty("customAttributes", "Working");

                var withoutAttributesOuterXml = withoutAttributes.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;
                var withAttributeOuterXml = withAttribute.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;

                Assert.AreEqual(withAttributeOuterXml, withoutAttributesOuterXml);
            }
        }

        [Test]
        public void SetAndUpdateCustomPropertyToDocument()
        {
            using (var withoutAttributes = new DocxDocument(Resources.DocumentWithoutAttributes))
            using (var withAttribute = new DocxDocument(Resources.DocumentWithAttribute))
            {
                withoutAttributes.SetCustomProperty("customAttributes", "Working1");
                withoutAttributes.SetCustomProperty("customAttributes", "Working");

                var withoutAttributesOuterXml = withoutAttributes.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;
                var withAttributeOuterXml = withAttribute.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;

                Assert.AreEqual(withAttributeOuterXml, withoutAttributesOuterXml);
            }
        }

        [Test]
        public void SetCustomPropertyToDocumentIfItAlreadyHasProperties()
        {
            using (var withTwoAttributes = new DocxDocument(Resources.DocumentWithTwoAttributes))
            using (var withAttribute = new DocxDocument(Resources.DocumentWithAttribute))
            {
                withAttribute.SetCustomProperty("customAttributes2", "Working2");

                var propertiesWithTwoAttributes = withTwoAttributes.GetWordDocument().CustomFilePropertiesPart.Properties;
                var propertiesWithAttribute = withAttribute.GetWordDocument().CustomFilePropertiesPart.Properties;

                Assert.AreEqual(propertiesWithTwoAttributes.First(x => x.LocalName == "property").OuterXml,
                                propertiesWithAttribute.First(x => x.LocalName == "property").OuterXml);
                Assert.AreEqual(propertiesWithTwoAttributes.Last(x => x.LocalName == "property").OuterXml,
                                propertiesWithAttribute.Last(x => x.LocalName == "property").OuterXml);
            }
        }

        [Test]
        public void GetCustomPropertyFromDocument()
        {
            using (var docxDocument = new DocxDocument(Resources.DocumentWithAttribute))
            {
                Assert.AreEqual("Working", docxDocument.GetCustomProperty("customAttributes"));
            }
        }

        [Test]
        public void GetCustomPropertyFromDocumentReturnNullIfItDoesNotExists()
        {
            using (var docxDocument = new DocxDocument(Resources.DocumentWithoutAttributes))
            {
                Assert.Null(docxDocument.GetCustomProperty("customAttributes"));
            }
        }

        [Test]
        public void AppendParagraphAddsNewParagraphToTheEndOfDocument()
        {
            using (var expected = new DocxDocument(Resources.DocumentWithAddedParagraph))
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                document.AppendParagraph("New paragraph content");

                Assert.AreEqual(GetParagraphs(expected).Count(), GetParagraphs(document).Count());
                Assert.AreEqual(GetParagraphs(expected).Last().InnerText, GetParagraphs(document).Last().InnerText);
            }
        }

        [Test]
        public void AppendHiddenParagraph()
        {
            using (var expected = new DocxDocument(Resources.DocumentWithAddedParagraph))
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
            	var paragraphsCount = GetParagraphs(document).Count;

                document.AppendParagraph("New paragraph content", false);

            	var paragraphs = GetParagraphs(document);
            	Assert.AreEqual(paragraphsCount + 1, paragraphs.Count);
				Assert.AreEqual(GetParagraphs(expected).Last().InnerXml, paragraphs.Last().InnerXml);
            }
        }

        [Test]
        public void AppendReplacementOpenCloseTagAddsTagsForReplacementsToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var initialCount = GetParagraphs(document).Count;

                document.AppendTag("Test");

                var paragraphs = GetParagraphs(document);

                Assert.AreEqual("{/Test}", paragraphs.Last().InnerText);
                Assert.AreEqual("{Test}", paragraphs.Skip(paragraphs.Count - 2).First().InnerText);
                Assert.AreEqual(initialCount + 2, paragraphs.Count);
            }
        }

        [Test]
        public void AppendSingleReplacementTagAddsTagForReplacementToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var initialCount = GetParagraphs(document).Count;

                document.AppendSelfclosingTag("Test");

                var paragraphs = GetParagraphs(document);

                Assert.AreEqual("{{Test}}", paragraphs.Last().InnerText);
                Assert.AreEqual(initialCount + 1, paragraphs.Count);
            }
        }

        [Test]
        public void AppendHiddenSingleReplacementTagAddsTagForReplacementToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var initialCount = GetParagraphs(document).Count;

                document.AppendSelfclosingTag("Test", false);

                var paragraphs = GetParagraphs(document);

                var insertedParagraph = paragraphs.Last();
                Assert.AreEqual("{{Test}}", insertedParagraph.InnerText);
                Assert.IsNotEmpty(insertedParagraph.ParagraphProperties);
                Assert.AreEqual(initialCount + 1, paragraphs.Count);
            }
        }

        private static ICollection<Paragraph> GetParagraphs(DocxDocument document)
        {
        	return document.GetWordDocument().MainDocumentPart.Document.Descendants().OfType<Paragraph>().ToArray();
        }

        [Test]
        public void HideContentInTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithVisibilityContentInTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithHideContentInTag))
            {
                initialDocument.SetTagVisibility("Tag", false);

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expectedDocument, initialDocument));
            }
        }

        [Test]
        public void VisibilityContentInTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithHideContentInTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithVisibilityContentInTag))
            {
                initialDocument.SetTagVisibility("Tag", true);

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expectedDocument, initialDocument));
            }
        }

        [Test]
        public void HidePlaceholderIfHideContentInTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithHideContentInPlaceholderInTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithVisibilityContentInPlaceholderInTag))
            {
                initialDocument.SetTagVisibility("Tag", true);

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expectedDocument, initialDocument));
            }
        }  
        
        [Test]
        public void NotVisibilityContentIfDocumentWithContentTypeTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithContentTypeTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithContentTypeTag))
            {
                initialDocument.SetTagVisibility("Tag", true);

                Assert.IsTrue(new DocxDocumentEqualityComparer().Equals(expectedDocument, initialDocument));
            }
        }

    	[Test]
    	public void CanCheckFieldAvailabilityUsingItsName()
    	{
			using (var document = new DocxDocument(Resources.WithSdtElements))
			{
				var hasBlockField = document.Fields.Contains("BlockField");
				var hasInlineField = document.Fields.Contains("InlineField");

				Assert.True(hasBlockField);
				Assert.True(hasInlineField);
			}
    	}
		
		[Test]
    	public void CanCheckFieldAvailabilityUsingItsTag()
    	{
			using (var document = new DocxDocument(Resources.WithSdtElements))
			{
				var hasBlockField = document.Fields.Contains(tag: "FirstTag");
				var hasInlineField = document.Fields.Contains(tag: "SecondTag");

				Assert.True(hasBlockField);
				Assert.True(hasInlineField);
			}
    	}

		[Test]
    	public void CanCorrectlyCountAllFieldsInDocument()
    	{
			using (var document = new DocxDocument(Resources.WithSdtElements))
			{
				Assert.AreEqual(2, document.Fields.Count());
			}
    	}

		[Test]
    	public void CanUnprotectNewDocument()
    	{
    		using (var document = new DocxDocument())
    		{
    			Assert.DoesNotThrow(document.Unprotect);
    		}
    	}

		[Test]
    	public void CanProtectNewDocument()
    	{
    		using (var document = new DocxDocument())
    		{
    			document.Protect();
    		}
    	}

		[Test]
		public void CanFindTagIfSdtElementHasNoSdtAlias()
		{
			using (var document = new DocxDocument(Resources.SdtElementWithoutSdtAlias))
			{
				document.Fields.Contains(tag: "FirstTag");
				document.Fields.Contains(tag: "SecondTag");
			}
		}
    }
}