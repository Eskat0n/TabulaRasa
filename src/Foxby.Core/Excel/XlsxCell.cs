namespace Foxby.Core.Excel
{
    public class XlsxCell : IXlsxCell
    {
        public string Content { get; set; }
        public IXlsxCellOption Option { get; set; }
    }
}