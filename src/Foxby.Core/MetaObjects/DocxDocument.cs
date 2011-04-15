using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace Foxby.Core.MetaObjects
{
	public class DocxDocument : IDisposable
	{
		private readonly MemoryStream documentStream;
		private readonly WordprocessingDocument wordDocument;

		public DocxDocument(byte[] template)
		{
			documentStream = new MemoryStream();
			documentStream.Write(template, 0, template.Length);

			wordDocument = WordprocessingDocument.Open(documentStream, true);
		}

		public void Dispose()
		{
			documentStream.Dispose();
		}

	    public byte[] ToArray()
		{
			wordDocument.Close();
			return documentStream.ToArray();
		}

	    public WordprocessingDocument GetWordDocument()
		{
			return wordDocument;
		}
	}
}