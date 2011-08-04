using System.Collections.Generic;

namespace Foxby.Core
{
    public interface IXlsxContent
    {
        void AddRow(IEnumerable<IXlsxCell> xlsxCell);
        IEnumerable<IEnumerable<IXlsxCell>> GetTable();
    }
}