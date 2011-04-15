using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder
{
	public class DocxDocumentPlaceholderContextBuilder : DocxDocumentContextBuilderBase
	{
		private readonly RunProperties runProperties;

		public DocxDocumentPlaceholderContextBuilder(WordprocessingDocument document, RunProperties runProperties)
			: base(document)
		{
			this.runProperties = runProperties;
		}

		protected override RunProperties RunProperties
		{
			get { return runProperties; }
		}
	}
}