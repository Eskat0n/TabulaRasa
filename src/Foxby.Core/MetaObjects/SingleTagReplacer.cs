namespace TabulaRasa.MetaObjects
{
    using DocumentFormat.OpenXml.Wordprocessing;

    internal class SingleTagReplacer : TagReplacer
    {
        public SingleTagReplacer(string name, DocxDocument document)
            : base(name, document)
        {
        }

        public override void Replace(string newValue)
        {
            Document.Replace(Name, newValue);
        }

        public override void Replace(Paragraph newValue)
        {
        }
    }
}