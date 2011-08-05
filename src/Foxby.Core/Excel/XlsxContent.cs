using System.Collections.Generic;

namespace Foxby.Core.Excel
{
    public class XlsxContent : IXlsxContent
    {
    	private readonly ICollection<IEnumerable<IXlsxCell>> table = new List<IEnumerable<IXlsxCell>>();

        public void AddRow(IEnumerable<IXlsxCell> xslCells)
        {
            table.Add(xslCells);
        }

    	public IEnumerable<IEnumerable<IXlsxCell>> Table
    	{
    		get { return table; }
    	}
    }
}