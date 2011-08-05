using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Foxby.Core.DocumentBuilder
{
	internal class DocxDocumentTableSchemeBuilder : DocxDocumentAggregationBuilder, IDocumentTableSchemeBuilder, IDocumentTableRowsBuilder
	{
		private readonly TableRow headerRow;
		private readonly Table table;

		internal DocxDocumentTableSchemeBuilder(WordprocessingDocument document, TableProperties contextTableProperties)
			: base(document)
		{
			table = new Table();

			if (contextTableProperties == null)
			{
				var borderType = new EnumValue<BorderValues>(BorderValues.Thick);
				var tblProp = new TableProperties(
					new TableBorders(
						new TopBorder {Val = borderType, Size = 1},
						new BottomBorder {Val = borderType, Size = 1},
						new LeftBorder {Val = borderType, Size = 1},
						new RightBorder {Val = borderType, Size = 1},
						new InsideHorizontalBorder {Val = borderType, Size = 1},
						new InsideVerticalBorder {Val = borderType, Size = 1}
						)
					);
				table.AppendChild(tblProp);
			}
			else
				table.AppendChild(contextTableProperties);

			headerRow = new TableRow();
			table.AppendChild(headerRow);
			Aggregation.Add(table);
		}

		public IDocumentTableRowsBuilder Row(params string[] content)
		{
			return Row(content.Select(AddText).ToArray());
		}

		public IDocumentTableRowsBuilder Row(params Action<ICellContextBuilder>[] options)
		{
			var row = new TableRow();
			table.Append(row);

			foreach (var option in options)
			{
				AddCell(row, option);
			}

			return this;
		}

		public IDocumentTableSchemeBuilder Column(string columnName)
		{
			AddCell(headerRow, AddText(columnName));

			return this;
		}

        public IDocumentTableSchemeBuilder Column(string columnName, float widthInPercents)
        {
            var width = Convert.ToInt32((widthInPercents*100)/15);
            var tableCellProperties = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = width.ToString() });
            AddCell(headerRow, AddText(columnName), tableCellProperties);

            return this;
        }

		private static Action<ICellContextBuilder> AddText(string text)
		{
			return x => x.AddText(text);
		}

        private void AddCell(OpenXmlCompositeElement row, Action<ICellContextBuilder> options)
        {
            var tableCellProperties = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto });
            AddCell(row, options, tableCellProperties);
        }

        private void AddCell(OpenXmlCompositeElement row, Action<ICellContextBuilder> options, TableCellProperties cellProperties)
		{
			var builder = new DocxDocumentCellContextBuilder(Document, cellProperties);

			options(builder);
			row.AppendChild(builder.ToElement());
		}
	}
}