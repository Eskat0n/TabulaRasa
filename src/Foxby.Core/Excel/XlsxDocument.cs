using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Foxby.Core.Excel
{
    ///<summary>
    /// 
    ///</summary>
    public class XlsxDocument : IDisposable
    {
        private readonly UInt32Value easyStyleIndex = 0U;        
        private readonly SheetData sheetData = new SheetData();
        private readonly MemoryStream workBookStream = new MemoryStream();
        private readonly SpreadsheetDocument xlsxWorkBook;
        private UInt32Value boldStyleIndex;
        private uint indexRow;

        public XlsxDocument()
        {
            xlsxWorkBook = SpreadsheetDocument.Create(workBookStream, SpreadsheetDocumentType.Workbook);

            WorkbookPart workBookPart = xlsxWorkBook.AddWorkbookPart();
            workBookPart.Workbook = new Workbook();

            var workSheetPart = workBookPart.AddNewPart<WorksheetPart>();
            workSheetPart.Worksheet = new Worksheet(sheetData);

            Sheets sheets = xlsxWorkBook.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet
                            {
                                Id = xlsxWorkBook.WorkbookPart
                                    .GetIdOfPart(workSheetPart),
                                SheetId = 1,
                                Name = "Лист1"
                            };
            sheets.Append(sheet);

            var workbookStylesPart = workBookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = CreateStylesheet();
        }

        public void Dispose()
        {
            workBookStream.Dispose();
        }

        public byte[] ToArray()
        {
            xlsxWorkBook.Close();
            return workBookStream.ToArray();
        }

        public void AddRow(IEnumerable<IXlsxCell> cels)
        {
            if (cels == null) throw new ArgumentNullException("cels");

            var row = new Row
            {
                RowIndex = ++indexRow
            };
            sheetData.Append(row);

            for (int i = 0; i < cels.Count(); i++)
            {
                UInt32Value styleIndex;
                if (cels.ElementAt(i).Option == null)
                    styleIndex = easyStyleIndex;
                else
                    styleIndex = cels.ElementAt(i).Option.Bold ? boldStyleIndex : easyStyleIndex;

                var cellValue = new CellValue(cels.ElementAt(i).Content);
                var cell = new Cell(cellValue)
                               {
                                   CellReference = "R" + indexRow + "C" + (i + 1),
                                   DataType = CellValues.String,
                                   StyleIndex = styleIndex
                               };

                row.Append(cell);
            }
        }

        private Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet();

            var fonts = new Fonts(new[]
                                      {
                                          new Font(),
                                          new Font(new Bold())
                                      });

            var fills = new Fills(new[] {new Fill()});
            var borders = new Borders(new[] {new Border()});


            OpenXmlElement boldFont = fonts.Where(x => x.ChildElements.Any(z => z is Bold)).Single();
            var boldFontId = new UInt32Value((uint) fonts.ToList().IndexOf(boldFont));

            var cellFormats = new CellFormats(new[]
                                                  {
                                                      new CellFormat(),
                                                      new CellFormat(),
                                                      new CellFormat {FontId = boldFontId}
                                                  });
            OpenXmlElement cellBoldFormat = cellFormats.Where(x => ((CellFormat) x).FontId == boldFontId).Single();
            boldStyleIndex = new UInt32Value((uint) cellFormats.ToList().IndexOf(cellBoldFormat));
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);

            stylesheet.Append(cellFormats);

            return stylesheet;
        }
    }
}