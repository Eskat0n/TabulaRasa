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
					          	.Indent.Paragraph("Тестовый 1")
					          	.Indent.Paragraph("Тестовый 2")
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
					     x => x.EditableStart().Indent.Paragraph("Тестовый 1")
					          	.EditableEnd()
					          	.Paragraph("Тестовый 2 нередактируемый"));

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
					          	.Indent.Paragraph("Тестовая строка 1\r\nтестовая строка 2", "тестовая строка 3")
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
					                                   	.Indent.Paragraph("Тестовый в новом теге 1")
					                                   	.Indent.Paragraph("Тестовый в новом теге 2 строка 1", "Тестовый в новом теге 2 строка 2")
					                                   	.EditableEnd())
					          	.EditableStart()
					          	.Indent.Paragraph("Тестовый 1")
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
					          	.Indent.Paragraph("Тестовый 1")
					          	.EditableEnd())
					.Tag("NEW_TAG",
					     x => x.EditableStart()
					          	.Indent.Paragraph("Тестовый в новом теге 1")
					          	.Indent.Paragraph("Тестовый в новом теге 2 строка 1", "Тестовый в новом теге 2 строка 2")
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
					.Placeholder("TITLE", x => x.Text("Заголовок"));

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
					          	.Paragraph("Ниже находятся три пустые строки")
					          	.EmptyLine()
					          	.EmptyLine()
					          	.EmptyLine()
					          	.Paragraph("Выше находятся три пустые строки")
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
					          	.OrderedList(z => z.Item("Элемент списка 1")
					          	                  	.Item("Элемент списка 2 строка 1", "Элемент списка 2 строка 2"))
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
					     x => x.Paragraph("Тестовый слева")
					          	.Center.Paragraph("Тестовый по центру")
					          	.Right.Paragraph("Тестовый справа")
					          	.Both.Paragraph("Тестовый по ширине"));

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
					     x => x.Paragraph(z => z.Text("Справа плейсхолдер ")
					                           	.Placeholder("SIMPLE")
					                           	.Text(" Слева плейсхолдер")));

				
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
					             x => x.Text(" Текст плейсхолдера "));

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
					.Tag("MAIN_CONTENT", x => x.Paragraph(z => z.Text("Справа плейсхолдер\r\n")
					                                           	.Placeholder("OUTER_PH", y => y.Text("Текст плейсхолдера справа внутренний\r\n")
					                                           	                              	.Placeholder("INNER_PH", m => m.Text("Внутренний текст\r\n"))
					                                           	                              	.Text("слева внутренний\r\n"))
					                                           	.Text("Слева плейсхолдер")));

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
					.Placeholder("TITLE", x => x.Text("Замененный текст"), false);

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

				builder.Tag("MAIN_CONTENT", x => x.Table(y => y.Column("Наименование")
				                                              	.Column("Адрес"),
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

				builder.Tag("MAIN_CONTENT", x => x.Table(y => y.Column("Наименование")
				                                              	.Column("Адрес"),
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

				builder.Tag("MAIN_CONTENT", x => x.Table(y => y.Column("Наименование")
				                                              	.Column("Адрес"),
				                                         y => y.Row("Нежилое помещение", "Челябинская область")
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

				builder.Tag("MAIN_CONTENT", x => x.BorderNone.Table(y => y.Column("Наименование")
				                                                         	.Column("Адрес"),
				                                                    y => y.Row("Нежилое помещение", "Челябинская область")
																		
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
					.Tag("SUB", x => x.Paragraph("Параграф во всех тегах"));

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
					.Placeholder("INNER", x => x.Text("Вставленный контент "));

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
					.Tag("NON_EXISTING", x => x.Paragraph("Тест")
					                          	.AppendTag("NEW", z => { })
					                          	.OrderedList(z => z.Item("Элемент 1").Item("Элемент 2"))
					                          	.EditableStart()
					                          	.Table(z => z.Column("Колонки"), z => z.Row("Строка"))
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
					                                  	.Text("Тест")
					                                  	.Placeholder("NEW")
					                                  	.Text("Тест")
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
					.Tag("MAIN_CONTENT", x => x.Paragraph(z => z.Text("Слово1").Text(" Слово2 ").Text(" Слово3")));

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
					.Tag("MAIN_CONTENT", x => x.OrderedList(z => z.Item(y => y.Text("Первый элемент списка").Text(" с пробелом"))
					                                             	.Item(y => y.EditableStart().Text("Редактируемый элемент списка").EditableEnd())
					                                             	.Item(y => y.Text("Справа плейсхолдер ").Placeholder("NEW").Text(" слева плейсхолдер"))));

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
					.Tag("MAIN_CONTENT", x => x.Paragraph(z => z.Bold.Text("Жирный ")
					                                           	.Italic.Text("Курсив ")
					                                           	.Underlined.Text("Подчеркнутый ")
					                                           	.Text("Нормальный")));

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
					.Tag("MAIN_CONTENT", x => x.Paragraph("Жирный ".Bold() + "Курсив ".Italic() + "Подчеркнутый ".Underlined() + "Нормальный"));

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
					.Placeholder("INNER", x => x.Bold.Text("Жирный ")
					                           	.Italic.Text("Курсив ")
					                           	.Underlined.Text("Подчеркнутый ")
					                           	.Text("Нормальный "));

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
					.Tag("MAIN_CONTENT", x => x.Table(z => z.Column("Первая").Column("Вторая"),
					                                  z => z.Row(y => y.Text("Первая строка в 1 колонке", "Вторая строка в 1 колонке"),
					                                             y => y.Text("Первая строка во 2 колонке").Text("\r\nВторая строка во 2 колонке"))));

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
					                          	.OrderedList(z => z.Item("Первый элемент списка").Item("Второй элемент списка"))
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
					                          	.OrderedList(z => z.Item("Первый элемент первого списка").Item("Второй элемент первого списка"))
					                          	.OrderedList(z => z.Item("Первый элемент второго списка").Item("Второй элемент второго списка"))
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