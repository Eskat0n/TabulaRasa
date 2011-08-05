using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.MetaObjects;
using Foxby.Core.Tests.DocumentBuilder;
using Foxby.Core.Tests.EqualityComparers;
using Foxby.Core.Tests.Properties;
using Xunit;
using Xunit.Sdk;

namespace Foxby.Core.Tests
{
    public class DocxDocumentTests
    {
        [Fact]
        public void DocumentsWithDifferentInnerXmlMustBeNotEqual()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            using (var otherDocument = new DocxDocument(Resources.DocumentWithParagraph))
            {
                var comparer = new DocxDocumentEqualityComparer();
                Assert.Throws<EqualException>(() => comparer.Equals(document, otherDocument));
            }
        }

        [Fact]
        public void DocumentsWithEqualInnerXmlMustBeEqual()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            using (var otherDocument = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                Assert.Equal(document, otherDocument, new DocxDocumentEqualityComparer());
            }
        }

        [Fact(Skip = "Несовпадающие rsidR у элементов параграфов документа")]
        public void TestCleanContent()
        {
            using (var withParagraph = new DocxDocument(Resources.DocumentWithParagraph))
            using (var withOutParagraph = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                withParagraph.CleanContent("Edit");
                Assert.Equal(withOutParagraph, withParagraph, new DocxDocumentEqualityComparer());
            }
        }

        [Fact]
        public void ReplacePairsFromDictionary()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithTitle))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithReplacedTitle))
            {
                initialDocument.Replace(new Dictionary<string, string>
				                        	{
				                        		{"Title", "Hello"}
				                        	});
                Assert.Equal(expectedDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }

        [Fact]
        public void ReplaceSingleTagWithTextBlocksDoesNothing()
        {
            using (var initialDocument = new DocxDocument(Resources.WithMainContentSingleTag))
            {
                var content = new[] { new TextBlock("Контент документа") };
                initialDocument.Replace("MAIN_CONTENT", content);

                Assert.Equal(initialDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }

        [Fact]
        public void ReplaceOpenCloseTagWithTextBlocksEditable()
        {
            using (var initialDocument = new DocxDocument(Resources.WithMainContentTag))
            using (var expectedDocument = new DocxDocument(Resources.WithMainContentInserted))
            {
                var content = new[] { new TextBlock("Контент документа") };
                initialDocument.Replace("MAIN_CONTENT", content);

                Assert.Equal(expectedDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }

        [Fact]
        public void InsertTagContentToOpenCloseTag()
        {
            using (var initialDocument = new DocxDocument(Resources.WithMainContentTag))
            using (var expectedDocument = new DocxDocument(Resources.WithMainContentInserted))
            {
                var content = new[] { new TextBlock("Контент документа") };
                initialDocument.InsertTagContent("MAIN_CONTENT", content);

                Assert.Equal(expectedDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }

        [Fact]
        public void UnprotectRemovesWrightProtectionFromFile()
        {
            using (var @protected = new DocxDocument(Resources.Protected))
            using (var @unprotected = new DocxDocument(Resources.Unprotected))
            {
                @protected.Unprotect();

                Assert.Equal(@unprotected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml, @protected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml);
            }
        }

        [Fact]
        public void ProtectAddWrightProtectionToFile()
        {
            using (var @protected = new DocxDocument(Resources.Protected))
            using (var @unprotected = new DocxDocument(Resources.Unprotected))
            {
                @unprotected.Protect();

                Assert.Equal(@protected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml, @unprotected.GetWordDocument().MainDocumentPart.DocumentSettingsPart.Settings.OuterXml);
            }
        }

        [Fact]
        public void SetCustomPropertyToDocument()
        {
            using (var withoutAttributes = new DocxDocument(Resources.DocumentWithoutAttributes))
            using (var withAttribute = new DocxDocument(Resources.DocumentWithAttribute))
            {
                withoutAttributes.SetCustomProperty("customAttributes", "Working");

                var withoutAttributesOuterXml = withoutAttributes.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;
                var withAttributeOuterXml = withAttribute.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;
                Assert.Equal(withAttributeOuterXml, withoutAttributesOuterXml);
            }
        }

        [Fact]
        public void SetAndUpdateCustomPropertyToDocument()
        {
            using (var withoutAttributes = new DocxDocument(Resources.DocumentWithoutAttributes))
            using (var withAttribute = new DocxDocument(Resources.DocumentWithAttribute))
            {
                withoutAttributes.SetCustomProperty("customAttributes", "Working1");
                withoutAttributes.SetCustomProperty("customAttributes", "Working");

                var withoutAttributesOuterXml = withoutAttributes.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;
                var withAttributeOuterXml = withAttribute.GetWordDocument().CustomFilePropertiesPart.Properties.Single(x => x.LocalName == "property").OuterXml;
                Assert.Equal(withAttributeOuterXml, withoutAttributesOuterXml);
            }
        }

        [Fact]
        public void SetCustomPropertyToDocumentIfItAlreadyHasProperties()
        {
            using (var withTwoAttributes = new DocxDocument(Resources.DocumentWithTwoAttributes))
            using (var withAttribute = new DocxDocument(Resources.DocumentWithAttribute))
            {
                withAttribute.SetCustomProperty("customAttributes2", "Working2");

                DocumentFormat.OpenXml.CustomProperties.Properties propertiesWithTwoAttributes = withTwoAttributes.GetWordDocument().CustomFilePropertiesPart.Properties;
                DocumentFormat.OpenXml.CustomProperties.Properties propertiesWithAttribute = withAttribute.GetWordDocument().CustomFilePropertiesPart.Properties;

                Assert.Equal(propertiesWithTwoAttributes.First(x => x.LocalName == "property").OuterXml,
                             propertiesWithAttribute.First(x => x.LocalName == "property").OuterXml);
                Assert.Equal(propertiesWithTwoAttributes.Last(x => x.LocalName == "property").OuterXml,
                             propertiesWithAttribute.Last(x => x.LocalName == "property").OuterXml);
            }
        }

        [Fact]
        public void GetCustomPropertyFromDocument()
        {
            using (var docxDocument = new DocxDocument(Resources.DocumentWithAttribute))
            {
                Assert.Equal("Working", docxDocument.GetCustomProperty("customAttributes"));
            }
        }

        [Fact]
        public void GetCustomPropertyFromDocumentReturnNullIfItDoesNotExists()
        {
            using (var docxDocument = new DocxDocument(Resources.DocumentWithoutAttributes))
            {
                Assert.Null(docxDocument.GetCustomProperty("customAttributes"));
            }
        }

        [Fact]
        public void AppendParagraphAddsNewParagraphToTheEndOfDocument()
        {
            using (var expected = new DocxDocument(Resources.DocumentWithAddedParagraph))
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                document.AppendParagraph("New paragraph content");

                Assert.Equal(GetParagraphs(expected).Count(), GetParagraphs(document).Count());
                Assert.Equal(GetParagraphs(expected).Last().InnerText, GetParagraphs(document).Last().InnerText);
            }
        }

        [Fact]
        public void AppendHiddenParagraph()
        {
            using (var expected = new DocxDocument(Resources.DocumentWithAddedParagraph))
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var paragraphsCount = GetParagraphs(document).Count();

                document.AppendParagraph("New paragraph content", false);

                Assert.Equal(paragraphsCount + 1, GetParagraphs(document).Count());
                Assert.Equal(GetParagraphs(expected).Last().InnerXml, GetParagraphs(document).Last().InnerXml);
            }
        }

        [Fact]
        public void AppendReplacementOpenCloseTagAddsTagsForReplacementsToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var initialCount = GetParagraphs(document).Count();

                document.AddOpenCloseTag("Test");

                var paragraphs = GetParagraphs(document);

                Assert.Equal("{/Test}", paragraphs.Last().InnerText);
                Assert.Equal("{Test}", paragraphs.Skip(paragraphs.Count() - 2).First().InnerText);
                Assert.Equal(initialCount + 2, paragraphs.Count());
            }
        }

        [Fact]
        public void AppendSingleReplacementTagAddsTagForReplacementToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var initialCount = GetParagraphs(document).Count();

                document.AddSingleTag("Test");

                var paragraphs = GetParagraphs(document);

                Assert.Equal("{{Test}}", paragraphs.Last().InnerText);
                Assert.Equal(initialCount + 1, paragraphs.Count());
            }
        }

        [Fact]
        public void AppendHiddenSingleReplacementTagAddsTagForReplacementToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var initialCount = GetParagraphs(document).Count();

                document.AddSingleTag("Test", false);

                var paragraphs = GetParagraphs(document);

                var insertedParagraph = paragraphs.Last();
                Assert.Equal("{{Test}}", insertedParagraph.InnerText);
                Assert.NotEmpty(insertedParagraph.ParagraphProperties);
                Assert.Equal(initialCount + 1, paragraphs.Count());
            }
        }

        [Fact]
        public void AppendParagraphWithTextBlocksAddsTextToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.DocumentWithoutParagraph))
            {
                var initialCount = GetParagraphs(document).Count();

                document.AppendParagraph(new[] { new TextBlock("text of the paragraph "), new TextBlock("another part") });

                var paragraphs = GetParagraphs(document);

                Assert.Equal(initialCount + 1, paragraphs.Count());
                Assert.Equal("text of the paragraph another part", paragraphs.Last().InnerText);
            }
        }

        [Fact]
        public void AppendParagraphWithProtectedTextAddsTextWithPermissionsRequireToTheEndOfTheDocument()
        {
            using (var document = new DocxDocument(Resources.Unprotected))
            {
                var initialCount = GetParagraphs(document).Count();

                document.AppendParagraph(new[]
				                         	{
				                         		new TextBlock("part 1 ", false), 
												new TextBlock("part 2 "),
												new TextBlock("part 3", false)
				                         	});

                var paragraphs = GetParagraphs(document);

                Assert.Equal(initialCount + 1, paragraphs.Count());
                Assert.Equal("part 1 part 2 part 3", paragraphs.Last().InnerText);
                Assert.Contains("permEnd", paragraphs.Last().ChildElements.Select(x => x.LocalName));
                Assert.Contains("permStart", paragraphs.Last().ChildElements.Select(x => x.LocalName));
                Assert.Equal("part 2 ", paragraphs.Last().ChildElements.Where(x => x.LocalName == "permStart").First().NextSibling().InnerText);
            }
        }

        private static IEnumerable<Paragraph> GetParagraphs(DocxDocument document)
        {
            return document.GetWordDocument().MainDocumentPart.Document.Descendants().OfType<Paragraph>();
        }

        [Fact]
        public void HideContentInTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithVisibilityContentInTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithHideContentInTag))
            {
                initialDocument.SetTagVisibility("Tag", false);
                Assert.Equal(expectedDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }

        [Fact]
        public void VisibilityContentInTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithHideContentInTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithVisibilityContentInTag))
            {
                initialDocument.SetTagVisibility("Tag", true);
                Assert.Equal(expectedDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }

        [Fact]
        public void HidePlaceholderIfHideContentInTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithHideContentInPlaceholderInTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithVisibilityContentInPlaceholderInTag))
            {
                initialDocument.SetTagVisibility("Tag", true);
                Assert.Equal(expectedDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }  
        
        [Fact]
        public void NotVisibilityContentIfDocumentWithContentTypeTag()
        {
            using (var initialDocument = new DocxDocument(Resources.DocumentWithContentTypeTag))
            using (var expectedDocument = new DocxDocument(Resources.DocumentWithContentTypeTag))
            {
                initialDocument.SetTagVisibility("Tag", true);
                Assert.Equal(expectedDocument, initialDocument, new DocxDocumentEqualityComparer());
            }
        }
    }
}