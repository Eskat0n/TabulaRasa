namespace Foxby.Core
{
    public interface IXlsxCell
    {
        string Content { get; set; }
        IXlsxCellOption Option { get; set; }
    }
}