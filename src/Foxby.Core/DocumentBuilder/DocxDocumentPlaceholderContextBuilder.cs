namespace TabulaRasa.DocumentBuilder
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal class DocxDocumentPlaceholderContextBuilder : DocxDocumentContextBuilderBase
	{
		private readonly RunProperties _runProperties;

		public DocxDocumentPlaceholderContextBuilder(WordprocessingDocument document, RunProperties runProperties)
			: base(document)
		{
			_runProperties = runProperties;
		}

		protected override RunProperties RunProperties
		{
			get { return _runProperties; }
		}
	}
}