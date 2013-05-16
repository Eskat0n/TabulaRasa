namespace TabulaRasa.Tests.DocumentBuilder
{
    using System.IO;
    using NUnit.Framework;
    using Properties;
    using TabulaRasa.DocumentBuilder;
    using MetaObjects;

    [TestFixture]
    [Ignore("Test incomplete")]
    public class DocxDocumentBuilderMediaTests
    {
        [Test]
        public void CanInsertJpegImageIntoParagraph()
        {
            using (var expected = new DocxDocument(Resources.WithMainContentTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
                var builder = DocxDocumentBuilder.Create(document);

                var content = File.ReadAllBytes("Resources/Images/apple.jpg");
			    builder
                    .Tag("MAIN_CONTENT", x => x.Paragraph(z => z.Image(content, "image/jpeg")));

                SaveDocxFile(document, "WithImageInserted");
			}
        }

        private static void SaveDocxFile(DocxDocument document, string fileName)
        {
            File.WriteAllBytes(string.Format(@"D:\{0}.docx", fileName), document.ToArray());
        }
    }
}