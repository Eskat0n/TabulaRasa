namespace TabulaRasa.Tests.DocumentBuilder
{
    using System.IO;
    using EqualityComparers;
    using TabulaRasa.Tests.Properties;
    using Xunit;
    using TabulaRasa.DocumentBuilder;
    using TabulaRasa.MetaObjects;
    using TabulaRasa.MetaObjects.Extensions;

    public class DocxDocumentBuilderTests
	{
	    private static DocxDocumentBuilder CreateBuilder(DocxDocument document)
	    {
            return new DocxDocumentBuilder(document, new TagVisibilityOptions("Black_White_Template", new[] { "sg", "fg" }));
	    }

		[Fact]
		public void GoingToTagCleansItsEntireContent()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentTag))
			using (var document = new DocxDocument(Resources.WithMainContentInserted))
			{
				var builder = CreateBuilder(document);

				builder.Tag("MAIN_CONTENT", x => { });

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertTwoEditableIndentedParagraphsIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentInsertedTwoParagraphs))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.EditableStart()
					          	.Indent.Paragraph("�������� 1")
					          	.Indent.Paragraph("�������� 2")
					          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertOneEditableAndOneNonEditableNonIndentedParagraphIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentTwoDifferentParagraphs))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.EditableStart().Indent.Paragraph("�������� 1")
					          	.EditableEnd()
					          	.Paragraph("�������� 2 ���������������"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertOneEditableMultilineParagraphIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentInsertedMultilineParagraph))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.EditableStart()
					          	.Indent.Paragraph("�������� ������ 1\r\n�������� ������ 2", "�������� ������ 3")
					          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertNewTagAndFillItContentInPlaceIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentAndNewTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.AppendTag("NEW_TAG", z => z.EditableStart()
					                                   	.Indent.Paragraph("�������� � ����� ���� 1")
					                                   	.Indent.Paragraph("�������� � ����� ���� 2 ������ 1", "�������� � ����� ���� 2 ������ 2")
					                                   	.EditableEnd())
					          	.EditableStart()
					          	.Indent.Paragraph("�������� 1")
					          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanAccessNewlyCreatedTagViaTagAndChangeItsContent()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentAndNewTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.AppendTag("NEW_TAG", z => { })
					          	.EditableStart()
					          	.Indent.Paragraph("�������� 1")
					          	.EditableEnd())
					.Tag("NEW_TAG",
					     x => x.EditableStart()
					          	.Indent.Paragraph("�������� � ����� ���� 1")
					          	.Indent.Paragraph("�������� � ����� ���� 2 ������ 1", "�������� � ����� ���� 2 ������ 2")
					          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanReplacePlaceholderWithText()
		{
			using (var expected = new DocxDocument(Resources.WithTitlePlaceholderReplaced))
			using (var document = new DocxDocument(Resources.WithTitlePlaceholder))
			{
				var builder = CreateBuilder(document);

				builder
					.Placeholder("TITLE", x => x.Text("���������"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertEditableEmptyLinesIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentAndThreeEmptyLines))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.EditableStart()
					          	.Paragraph("���� ��������� ��� ������ ������")
					          	.EmptyLine()
					          	.EmptyLine()
					          	.EmptyLine()
					          	.Paragraph("���� ��������� ��� ������ ������")
					          	.EditableEnd());
		
				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertNumberOfEmptyLines()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
                var builder = CreateBuilder(expected);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.EmptyLine()
					          	.EmptyLine()
					          	.EmptyLine());

                builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.EmptyLine(3));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertEditableOrderedListIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentInsertedOrderedList))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.EditableStart()
					          	.OrderedList(z => z.Item("������� ������ 1")
					          	                  	.Item("������� ������ 2 ������ 1", "������� ������ 2 ������ 2"))
					          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertThreeParagraphsWithDifferentJustificationIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentInsertedJustifiedParagraphs))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.Paragraph("�������� �����")
					          	.Center.Paragraph("�������� �� ������")
					          	.Right.Paragraph("�������� ������")
					          	.Both.Paragraph("�������� �� ������"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertPlaceholderIntoParagraph()
		{
			using (var expected = new DocxDocument(Resources.WithPlaceholderInsertedInParagraph))
			using (var document = new DocxDocument(Resources.WithTitlePlaceholder))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT",
					     x => x.Paragraph(z => z.Text("������ ����������� ")
					                           	.Placeholder("SIMPLE")
					                           	.Text(" ����� �����������")));

				
				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertTextContentIntoInsertedPlaceholder()
		{
			using (var expected = new DocxDocument(Resources.WithPlaceholderInsertedContentInserted))
			using (var document = new DocxDocument(Resources.WithPlaceholderInsertedInParagraph))
			{
				var builder = CreateBuilder(document);

				builder
					.Placeholder("SIMPLE",
					             x => x.Text(" ����� ������������ "));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertPlaceholderIntoPlaceholderAndReplaceItsContent()
		{
			using (var expected = new DocxDocument(Resources.WithPlaceholderInPlaceholder))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.Paragraph(z => z.Text("������ �����������\r\n")
					                                           	.Placeholder("OUTER_PH", y => y.Text("����� ������������ ������ ����������\r\n")
					                                           	                              	.Placeholder("INNER_PH", m => m.Text("���������� �����\r\n"))
					                                           	                              	.Text("����� ����������\r\n"))
					                                           	.Text("����� �����������")));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertContentIntoPlaceholderAndRemoveIt()
		{
			using (var expected = new DocxDocument(Resources.WithTitlePlaceholderRemoved))
			using (var document = new DocxDocument(Resources.WithTitlePlaceholder))
			{
				var builder = CreateBuilder(document);

				builder
					.Placeholder("TITLE", x => x.Text("���������� �����"), false);

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertTableWithDifferentJustificationToCells()
		{
			using (var expected = new DocxDocument(Resources.TableWithFormattedCells))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder.Tag("MAIN_CONTENT", x => x.Table(y => y.Column("������������")
				                                              	.Column("�����"),
				                                         y => y.Row(z=>z.Left.Cell("123"), z=>z.Right.Cell("456"))
				                                 	));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertActionToCells()
		{
			using (var expected = new DocxDocument(Resources.WithCellWithPlaceholder))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder.Tag("MAIN_CONTENT", x => x.Table(y => y.Column("������������")
				                                              	.Column("�����"),
				                                         y => y.Row(z => z.Left.Cell(c => c.Placeholder("somePlaceholder", "hello world!")))
				                                 	));
				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertTableInOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithTableInsert))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder.Tag("MAIN_CONTENT", x => x.Table(y => y.Column("������������")
				                                              	.Column("�����"),
				                                         y => y.Row("������� ���������", "����������� �������")
				                                 	));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertTableWithoutBordersInOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithTableWithoutBordersInsert))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder.Tag("MAIN_CONTENT", x => x.BorderNone.Table(y => y.Column("������������")
				                                                         	.Column("�����"),
				                                                    y => y.Row("������� ���������", "����������� �������")
																		
				                                 	));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertContentInManyTagsWithSameName()
		{
			using (var expected = new DocxDocument(Resources.WithManyTagsInsertedParagraph))
			using (var document = new DocxDocument(Resources.WithManyTags))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("SUB", x => x.Paragraph("�������� �� ���� �����"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanReplaceManyPlaceholdersWithContent()
		{
			using (var expected = new DocxDocument(Resources.WithManyPlaceholdersInsertedContent))
			using (var document = new DocxDocument(Resources.WithManyPlaceholders))
			{
				var builder = CreateBuilder(document);

				builder
					.Placeholder("INNER", x => x.Text("����������� ������� "));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CallToNonExistingTagDoesNotChangeDocument()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("NON_EXISTING", x => x.Paragraph("����")
					                          	.AppendTag("NEW", z => { })
					                          	.OrderedList(z => z.Item("������� 1").Item("������� 2"))
					                          	.EditableStart()
					                          	.Table(z => z.Column("�������"), z => z.Row("������"))
					                          	.Paragraph(z => z.Placeholder("TEST"))
					                          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CallToNonExistingPlaceholderDoesNotChangeDocument()
		{
			using (var expected = new DocxDocument(Resources.WithTitlePlaceholder))
			using (var document = new DocxDocument(Resources.WithTitlePlaceholder))
			{
				var builder = CreateBuilder(document);

				builder
					.Placeholder("NON_EXISTING", x => x.EditableStart()
					                                  	.Text("����")
					                                  	.Placeholder("NEW")
					                                  	.Text("����")
					                                  	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertTextIntoParagraphWithTrailingSpacesPreserved()
		{
			using (var expected = new DocxDocument(Resources.WithMainContentInsertedTextWithSpaces))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.Paragraph(z => z.Text("�����1").Text(" �����2 ").Text(" �����3")));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertBuildedParagraphIntoOrderedListItem()
		{
			using (var expected = new DocxDocument(Resources.WithOrderedListWithParagraphs))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.OrderedList(z => z.Item(y => y.Text("������ ������� ������").Text(" � ��������"))
					                                             	.Item(y => y.EditableStart().Text("������������� ������� ������").EditableEnd())
					                                             	.Item(y => y.Text("������ ����������� ").Placeholder("NEW").Text(" ����� �����������"))));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void OpeningDocumentViaBuilderNormalizesItsPlaceholderRuns()
		{
			using (var expected = new DocxDocument(Resources.WithPlaceholdersNormalized))
			using (var document = new DocxDocument(Resources.WithPlaceholdersDenormalized))
			{
                CreateBuilder(document);

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertDifferentlyFormattedTextInsidePargraphIntoOpenCloseTag()
		{
			using (var expected = new DocxDocument(Resources.WithDifferentlyFormattedTextInTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.Paragraph(z => z.Bold.Text("������ ")
					                                           	.Italic.Text("������ ")
					                                           	.Underlined.Text("������������ ")
					                                           	.Text("����������")));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}
		
		[Fact]
		public void CanInsertDifferentlyFormattedTextInsidePargraphIntoOpenCloseTagViaFormat()
		{
			using (var expected = new DocxDocument(Resources.WithDifferentlyFormattedTextInTag))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.Paragraph("������ ".Bold() + "������ ".Italic() + "������������ ".Underlined() + "����������"));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertDifferentlyFormattedTextIntoPlaceholder()
		{
			using (var expected = new DocxDocument(Resources.WithDifferentlyFormattedTextInPlaceholder))
			using (var document = new DocxDocument(Resources.WithManyPlaceholders))
			{
				var builder = CreateBuilder(document);

				builder
					.Placeholder("INNER", x => x.Bold.Text("������ ")
					                           	.Italic.Text("������ ")
					                           	.Underlined.Text("������������ ")
					                           	.Text("���������� "));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertMultilineTextInTableCell()
		{
			using (var expected = new DocxDocument(Resources.WithTableWithMultilineCells))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.Table(z => z.Column("������").Column("������"),
					                                  z => z.Row(y => y.Text("������ ������ � 1 �������", "������ ������ � 1 �������"),
					                                             y => y.Text("������ ������ �� 2 �������").Text("\r\n������ ������ �� 2 �������"))));

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertIndentedOrderedListIntoTag()
		{
			using (var expected = new DocxDocument(Resources.WithIndentedOrderedListInserted))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.EditableStart()
					                          	.OrderedList(z => z.Item("������ ������� ������").Item("������ ������� ������"))
					                          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void CanInsertTwoIndependentOrderedListsIntoTag()
		{
			using (var expected = new DocxDocument(Resources.WithTwoIndependentOrderedListsInserted))
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				builder
					.Tag("MAIN_CONTENT", x => x.EditableStart()
					                          	.OrderedList(z => z.Item("������ ������� ������� ������").Item("������ ������� ������� ������"))
					                          	.OrderedList(z => z.Item("������ ������� ������� ������").Item("������ ������� ������� ������"))
					                          	.EditableEnd());

				Assert.Equal(expected, document, new DocxDocumentEqualityComparer());
			}
		}

		[Fact]
		public void ValidationForValidDocumentShouldBeCorrect()
		{
			using (var document = new DocxDocument(Resources.WithMainContentTag))
			{
				var builder = CreateBuilder(document);

				Assert.True(builder.Validate());
			}
		}

		[Fact]
		public void ValidationForInvalidDocumentShouldFail()
		{
			using (var document = new DocxDocument(Resources.InvalidDocument))
			{
				var builder = CreateBuilder(document);

				Assert.False(builder.Validate());
			}
		}

	    private static void SaveDocxFile(DocxDocument document, string fileName)
		{
			File.WriteAllBytes(string.Format(@"D:\{0}.docx", fileName), document.ToArray());
		}
	}
}