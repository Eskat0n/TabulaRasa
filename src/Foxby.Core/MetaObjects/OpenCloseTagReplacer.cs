using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.MetaObjects
{
    public class OpenCloseTagReplacer : TagReplacer
    {
        public OpenCloseTagReplacer(string name, DocxDocument document)
            : base(name, document)
        {
        }

        public override void Replace(string newValue)
        {
            Document.CleanContent(Name);

            Document.InsertTagContent(Name, new Paragraph(new Run(new Text(newValue))));
        }

        public override void Replace(Paragraph newValue)
        {
            Document.CleanContent(Name);

            Document.InsertTagContent(Name, newValue);
        }
    }
}