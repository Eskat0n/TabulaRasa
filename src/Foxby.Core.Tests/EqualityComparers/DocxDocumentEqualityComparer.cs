using System.Collections.Generic;
using Foxby.Core.MetaObjects;
using Xunit;

namespace Foxby.Core.Tests.EqualityComparers
{
	public class DocxDocumentEqualityComparer : IEqualityComparer<DocxDocument>
	{
		public bool Equals(DocxDocument x, DocxDocument y)
		{
			Assert.Equal(x.GetWordDocument().MainDocumentPart.Document.InnerXml, y.GetWordDocument().MainDocumentPart.Document.InnerXml);
			return true;
		}

		public int GetHashCode(DocxDocument obj)
		{
			return obj.GetWordDocument().MainDocumentPart.Document.InnerXml.GetHashCode();
		}
	}
}