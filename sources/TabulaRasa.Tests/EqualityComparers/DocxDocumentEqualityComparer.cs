namespace TabulaRasa.Tests.EqualityComparers
{
    using System.Collections.Generic;
    using System.Linq;
    using Xunit;
    using TabulaRasa.DocumentBuilder;
    using TabulaRasa.MetaObjects;

    public class DocxDocumentEqualityComparer : IEqualityComparer<DocxDocument>
	{
		public bool Equals(DocxDocument x, DocxDocument y)
		{
		    var enumerable1 = x.GetWordDocument().MainDocumentPart.GetRootElements().Select(e => e.InnerXml);
		    var enumerable2 = y.GetWordDocument().MainDocumentPart.GetRootElements().Select(e => e.InnerXml);
		    Assert.Equal(enumerable1, enumerable2);
			return true;
		}

		public int GetHashCode(DocxDocument obj)
		{
            return obj.GetWordDocument().MainDocumentPart.RootElement.InnerXml.GetHashCode();
		}
	}
}