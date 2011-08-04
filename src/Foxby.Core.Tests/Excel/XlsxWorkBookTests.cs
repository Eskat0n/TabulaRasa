using System;
using Foxby.Core.Excel;
using Xunit;

namespace Foxby.Core.Tests.Excel
{
    public class XlsxWorkBookTests
    {
        [Fact]
        public void IfCellsNullThenThowArgumentNullException()
        {
            var xlsxWorkBook = new XlsxDocument();

            Assert.Throws<ArgumentNullException>(() => xlsxWorkBook.AddRow(null));
        }
    }
}