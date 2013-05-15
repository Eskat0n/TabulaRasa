namespace TabulaRasa.Tests.DocumentBuilder
{
    using System.IO;
    using TabulaRasa.Tests.Properties;
    using Xunit;
    using TabulaRasa.DocumentBuilder;
    using TabulaRasa.MetaObjects;

    public class DocxDocumentBuilderMediaTests
    {
        [Fact]
        public void CanInsertJpegImageIntoParagraph()
        {
            using (var expected = new DocxDocument(Resources.WithMainContentTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
                var builder = DocxDocumentBuilder.Create(document);

                var content = File.ReadAllBytes("Resources/Images/apple.jpg");
			    builder
			        .Tag("MAIN_CONTENT",
			             x => x.Paragraph(z => z.Image(content, "image/jpeg")));

                SaveDocxFile(document, "WithImageInserted");
			}
        }

        private static void SaveDocxFile(DocxDocument document, string fileName)
        {
            File.WriteAllBytes(string.Format(@"D:\{0}.docx", fileName), document.ToArray());
        }
    }
}