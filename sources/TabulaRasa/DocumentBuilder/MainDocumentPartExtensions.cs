namespace TabulaRasa.DocumentBuilder
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;

    internal static class MainDocumentPartExtensions
    {
        public static IEnumerable<OpenXmlPartRootElement> GetRootElements(this MainDocumentPart mainPart)
        {
            yield return mainPart.RootElement;
            foreach (var part in mainPart.HeaderParts)
                yield return part.RootElement;
            foreach (var part in mainPart.FooterParts)
                yield return part.RootElement;
        }
    }
}