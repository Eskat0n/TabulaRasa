using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder;
using Foxby.Core.DocumentBuilder.Anchors;

namespace Foxby.Core.MetaObjects
{
	public class DocxDocument : IDisposable
	{
		private readonly MemoryStream documentStream;
		private readonly WordprocessingDocument wordDocument;

		public DocxDocument(byte[] template)
		{
			documentStream = new MemoryStream();
			documentStream.Write(template, 0, template.Length);

			wordDocument = WordprocessingDocument.Open(documentStream, true);
		}

		public bool HasSignatures
		{
			get
			{
			    DigitalSignatureOriginPart part = wordDocument.DigitalSignatureOriginPart;
			    return part != null && part.XmlSignatureParts != null && part.XmlSignatureParts.Any();
			}
		}

        public void UseTheme(TagVisibilityOptions theme)
        {
            SetVisibilityTag(theme.VisibleTagName, true);

            foreach (var notUsingTagName in theme.HiddenTagNames)
                SetVisibilityTag(notUsingTagName, false);
        }

	    public void SetVisibilityTag(string tagName, bool visible)
        {
            var documentTags = DocumentTag.Get(wordDocument, tagName);

            foreach (var documentTag in documentTags)
            {
                var paragraph = documentTag.Opening.NextSibling();
                while (paragraph as Paragraph != documentTag.Closing)
                {
                    SetVisibilityInParagraphProperty(paragraph, visible);

                    SetVisibilityInParagraphRuns(paragraph, visible);

                    paragraph = paragraph.NextSibling();
                }
            }
        }

	    private static void SetVisibilityInParagraphRuns(OpenXmlElement paragraph, bool visible)
	    {
	        var runsInParagraph = paragraph.Descendants().OfType<Run>().ToList();

	        foreach (var run in runsInParagraph)
	        {
	            if (IsPlaceholder(run))
	            {
	                Hide(run);
	                continue;
	            }

	            if (visible)
	                run.RunProperties.Vanish = null;
	            else if (run.RunProperties.Vanish == null)
	                Hide(run);
	        }
	    }

	    private static void SetVisibilityInParagraphProperty(OpenXmlElement paragraph, bool visible)
	    {
	        var paragraphProperties = paragraph.Descendants().OfType<ParagraphProperties>();
	        foreach (var paragraphProperty in paragraphProperties)
	        {
	            var paragraphMarkRunProperties = paragraphProperty.Descendants().OfType<ParagraphMarkRunProperties>().ToList();
	            if (paragraphMarkRunProperties.Any())
	                foreach (var markRunProperty in paragraphMarkRunProperties)
	                {
	                    markRunProperty.RemoveAllChildren<Vanish>();
	                    if (!visible)
	                        markRunProperty.Append(new Vanish());
	                }
	        }
	    }

	    private static void Hide(Run run)
	    {
	        run.RunProperties.Vanish = new Vanish();
	    }

	    private static bool IsPlaceholder(Run run)
	    {
	        return Regex.IsMatch(run.InnerText, @"^{{.*}}$") && run.RunProperties.Vanish != null;
	    }

	    public void Dispose()
		{
			documentStream.Dispose();
		}

		public void CleanContent(string tagName)
		{
			OpenXmlElement startTag = GetParagraph(GetOpenTagName(tagName));

			if (startTag == null) return;

			OpenXmlElement openXmlElement = startTag.NextSibling();
			while (openXmlElement.InnerText != GetCloseTagName(tagName))
			{
				openXmlElement.Remove();
				openXmlElement = startTag.NextSibling();
			}
			wordDocument.MainDocumentPart.Document.Save();
		}

		public void Replace(string singleTagName, string newValue)
		{
			string formattedReplacementName = GetSingleTagName(singleTagName);
		    var document = wordDocument.MainDocumentPart.Document;
		    document.InnerXml = document.InnerXml.Replace(formattedReplacementName, newValue);
            document.Save();
		}

		public void Replace(string tagName, IEnumerable<TextBlock> content)
		{
			var tagReplacer = GetTagReplacer(tagName);

			if (tagReplacer != null)
			{
				var paragraphContent = content.SelectMany(WrapText);
				var paragraph = new Paragraph(paragraphContent);

				tagReplacer.Replace(paragraph);
			}

			wordDocument.MainDocumentPart.Document.Save();
		}

		public void Replace(IEnumerable<KeyValuePair<string, string>> replacements)
		{
			foreach (var replacement in replacements)
			{
				TagReplacer tagReplacer = GetTagReplacer(replacement.Key);

				if (tagReplacer != null) 
					tagReplacer.Replace(replacement.Value);
			}

			wordDocument.MainDocumentPart.Document.Save();
		}

	    private OpenXmlElement GetParagraph(string formattedName)
		{
			IEnumerable<Paragraph> paragraphs = wordDocument.MainDocumentPart.Document.Descendants().OfType<Paragraph>();

			return paragraphs.SingleOrDefault(x => x.InnerText == formattedName);
		}

		private TagReplacer GetTagReplacer(string name)
		{
			if(ExistsUniqueTagWithInnerText(GetSingleTagName(name)))
			{
				return new SingleTagReplacer(name, this);
			}

			if (ExistsUniqueTagWithInnerText(GetOpenTagName(name)) && ExistsUniqueTagWithInnerText(GetCloseTagName(name)))
			{
				return new OpenCloseTagReplacer(name, this);
			}

			return null;
		}

		private bool ExistsUniqueTagWithInnerText(string text)
		{
			return wordDocument.MainDocumentPart.Document.Descendants().Count(x => x.InnerXml == text) == 1;
		}

		public byte[] ToArray()
		{
			wordDocument.Close();
			return documentStream.ToArray();
		}

		public void Unprotect()
		{
			SetProtectionAttribute("None");
		}

		public void Protect()
		{
			SetProtectionAttribute("readOnly");
		}

		private void SetProtectionAttribute(string value)
		{
			Settings settings = wordDocument.MainDocumentPart.DocumentSettingsPart.Settings;
			OpenXmlElement protectionTag = settings.Where(x => x.LocalName == "documentProtection").FirstOrDefault();

			if (protectionTag != null)
				protectionTag.SetAttribute(new OpenXmlAttribute("w:edit", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", value));
		}

		private static string GetOpenTagName(string tagName)
		{
			return string.Format("{{{0}}}", tagName);
		}

		private static string GetCloseTagName(string tagName)
		{
			return string.Format("{{/{0}}}", tagName);
		}

		private static string GetSingleTagName(string replacementName)
		{
			return string.Format("{{{{{0}}}}}", replacementName);
		}

		public void SetCustomProperty(string name, string value)
		{
			if (wordDocument.CustomFilePropertiesPart == null)
			{
				CustomFilePropertiesPart addCustomFilePropertiesPart = wordDocument.AddCustomFilePropertiesPart();
				addCustomFilePropertiesPart.Properties = new Properties();
			}

			Properties properties = wordDocument.CustomFilePropertiesPart.Properties;

			var existsProperty = properties.SingleOrDefault(p => ((CustomDocumentProperty)p).Name.Value == name);
			if(existsProperty != null)
			{
				existsProperty.Remove();
			}

			Int32Value nextPropertyId = properties
			                            	.Elements<CustomDocumentProperty>()
			                            	.Select(x => x.PropertyId.Value)
			                            	.Union(new[] {1})
			                            	.Max() + 1;

			properties.Append(new CustomDocumentProperty(new VTLPWSTR(value))
			                                                        	{
			                                                        		FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
			                                                        		PropertyId = nextPropertyId,
			                                                        		Name = name
			                                                        	});
		}

		public string GetCustomProperty(string name)
		{
			if (wordDocument.CustomFilePropertiesPart == null)
				return null;

			var openXmlElement = wordDocument.CustomFilePropertiesPart.Properties.FirstOrDefault(x => x.GetAttribute("name", string.Empty).Value == name);

			if (openXmlElement == null)
				return null;

			return openXmlElement.InnerText;
		}

		public void AppendParagraph(IEnumerable<TextBlock> content)
		{
			var paragraphContent = content.SelectMany(WrapText);
			var paragraph = new Paragraph(paragraphContent);

			wordDocument.MainDocumentPart.Document.Body.Append(paragraph);
			wordDocument.MainDocumentPart.Document.Save();
		}

		public void AppendParagraph(string content, bool visible = true)
		{
			var paragraph = new Paragraph();
			var run = new Run(new Text(content));
			
			if (visible == false)
			{
				run.RunProperties = new RunProperties(new Vanish());
				paragraph.ParagraphProperties = new ParagraphProperties(new RunProperties(new Vanish()));
			}

			paragraph.AppendChild(run);

			wordDocument.MainDocumentPart.Document.Body.Append(paragraph);
			wordDocument.MainDocumentPart.Document.Save();
		}

		public void InsertTagContent(string tagName, IEnumerable<TextBlock> contentForInsert)
		{
			var paragraphContent = contentForInsert.SelectMany(WrapText);
			var paragraph = new Paragraph(paragraphContent);

			InsertTagContent(tagName, paragraph);
		}

		public void InsertTagContent(string tagName, OpenXmlElement forInsert)
		{
			var tag = GetParagraph(GetCloseTagName(tagName));

			if (tag != null)
				tag.InsertBeforeSelf(forInsert);

			wordDocument.MainDocumentPart.Document.Save();
		}

	    internal WordprocessingDocument GetWordDocument()
		{
			return wordDocument;
		}

		public void AddOpenCloseTag(string name)
		{
			AppendParagraph(GetOpenTagName(name), false);
			AppendParagraph(GetCloseTagName(name), false);
		}

		public void AddSingleTag(string name, bool visible = true)
		{
			AppendParagraph(GetSingleTagName(name), visible);
		}

		private IEnumerable<OpenXmlElement> WrapText(TextBlock textBlock)
		{
			OpenXmlElement content;
			if (textBlock.Text == "\t") content = new TabChar();
			else content = new Text(textBlock.Text);

			var text = new Run(content);

			if (textBlock.Editable)
			{
				var id = wordDocument.MainDocumentPart.Document.Descendants().OfType<PermStart>()
					.Select(x => x.Id.Value)
					.Union(new[] {1})
					.Max();

				var permStart = new PermStart
				                	{
				                		Id = id,
				                		EditorGroup = RangePermissionEditingGroupValues.Everyone
				                	};

				var permEnd = new PermEnd {Id = id};

				return new OpenXmlElement[] {permStart, text, permEnd};
			}

			return new OpenXmlElement[] {text};
		}
	}
}