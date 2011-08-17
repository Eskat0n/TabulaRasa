using System.IO;
using Foxby.Core.DocumentBuilder;
using Foxby.Core.MetaObjects;
using Foxby.Core.Tests.EqualityComparers;
using Foxby.Core.Tests.Properties;
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

		private static void SaveDocxFile(DocxDocument document, string fileName)
		{
			File.WriteAllBytes(string.Format(@"D:\{0}.docx", fileName), document.ToArray());
		}
	}
}