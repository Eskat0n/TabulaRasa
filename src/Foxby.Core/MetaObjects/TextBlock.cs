namespace Foxby.Core.MetaObjects
{
    public class TextBlock
    {
        public bool Editable { get; private set; }

        public string Text { get; private set; }

        public TextBlock(string text, bool editable = true)
        {
            Text = text;
            Editable = editable;
        }
    }
}