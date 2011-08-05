using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.MetaObjects
{
	internal abstract class TagReplacer
    {
        protected readonly DocxDocument Document;

        protected TagReplacer(string name, DocxDocument document)
        {
            Document = document;
            Name = name;
        }

        protected string Name { get; private set; }

        public abstract void Replace(string newValue);

        public abstract void Replace(Paragraph newValue);
    }
}