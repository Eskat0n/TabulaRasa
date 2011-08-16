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
using Foxby.Core.MetaObjects.Collections;

namespace Foxby.Core.MetaObjects
{
	///<summary>
	/// Wrapper for OpenXML docx file which provides base methods for operating document
	///</summary>
	public class DocxDocument : IDisposable
	{
		private readonly MemoryStream _documentStream;
		private readonly WordprocessingDocument _wordDocument;
		private readonly FieldsCollection _fields;

		///<summary>
		/// Creates new instance of DocxDocument from scratch
		///</summary>
		public DocxDocument()
		{
			_documentStream = new MemoryStream();
		
			_wordDocument = WordprocessingDocument.Create(_documentStream, WordprocessingDocumentType.Document, true);
			_wordDocument.AddMainDocumentPart();
			_wordDocument.MainDocumentPart.Document = new Document(new Body());

			var sdtElements = _wordDocument.MainDocumentPart.Document.Body.Descendants<SdtElement>();
			_fields = new FieldsCollection(sdtElements);
		}

		///<summary>
		/// Creates new instance of DocxDocument using <paramref name="template"/>
		///</summary>
		///<param name="template">Binary content of docx file</param>
		public DocxDocument(byte[] template)
		{
			_documentStream = new MemoryStream();
			_documentStream.Write(template, 0, template.Length);

			_wordDocument = WordprocessingDocument.Open(_documentStream, true);

			var sdtElements = _wordDocument.MainDocumentPart.Document.Body.Descendants<SdtElement>();
			_fields = new FieldsCollection(sdtElements);
		}

		///<summary>
		/// Checks whether docx document have any digital signatures or not
		///</summary>
		public bool HasSignatures
		{
			get
			{
			    DigitalSignatureOriginPart part = _wordDocument.DigitalSignatureOriginPart;
			    return part != null && part.XmlSignatureParts != null && part.XmlSignatureParts.Any();
			}
		}

		/// <summary>
		/// Sets visibility as specified in <paramref name="options"/> passed
		/// </summary>
		/// <param name="options">Specify tags to be shown and hidden</param>
        public void SetTagVisibility(TagVisibilityOptions options)
        {
            SetTagVisibility(options.VisibleTagName, true);

            foreach (var notUsingTagName in options.HiddenTagNames)
                SetTagVisibility(notUsingTagName, false);
        }

	    ///<summary>
	    /// Sets visibility for tag specified by <paramref name="tagName"/>
	    ///</summary>
	    ///<param name="tagName">Tag name</param>
	    ///<param name="isVisible">Shows tag if set to true, otherwise hides it</param>
	    public void SetTagVisibility(string tagName, bool isVisible)
        {
            var documentTags = DocumentTag.Get(_wordDocument, tagName);

            foreach (var documentTag in documentTags)
            {
                var paragraph = documentTag.Opening.NextSibling();
                while (paragraph as Paragraph != documentTag.Closing)
                {
                    SetVisibilityInParagraphProperty(paragraph, isVisible);

                    SetVisibilityInParagraphRuns(paragraph, isVisible);

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
			_documentStream.Dispose();
		}

		///<summary>
		/// Removes all content from tag specified by <paramref name="tagName"/>
		///</summary>
		///<param name="tagName">Tag name</param>
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
			_wordDocument.MainDocumentPart.Document.Save();
		}

		internal void Replace(string singleTagName, string newValue)
		{
			string formattedReplacementName = GetSingleTagName(singleTagName);
		    var document = _wordDocument.MainDocumentPart.Document;
		    document.InnerXml = document.InnerXml.Replace(formattedReplacementName, newValue);
            document.Save();
		}

		internal void Replace(string tagName, IEnumerable<TextBlock> content)
		{
			var tagReplacer = GetTagReplacer(tagName);

			if (tagReplacer != null)
			{
				var paragraphContent = content.SelectMany(WrapText);
				var paragraph = new Paragraph(paragraphContent);

				tagReplacer.Replace(paragraph);
			}

			_wordDocument.MainDocumentPart.Document.Save();
		}

		internal void Replace(IEnumerable<KeyValuePair<string, string>> replacements)
		{
			foreach (var replacement in replacements)
			{
				TagReplacer tagReplacer = GetTagReplacer(replacement.Key);

				if (tagReplacer != null) 
					tagReplacer.Replace(replacement.Value);
			}

			_wordDocument.MainDocumentPart.Document.Save();
		}

	    private OpenXmlElement GetParagraph(string formattedName)
		{
			IEnumerable<Paragraph> paragraphs = _wordDocument.MainDocumentPart.Document.Descendants().OfType<Paragraph>();

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
			return _wordDocument.MainDocumentPart.Document.Descendants().Count(x => x.InnerXml == text) == 1;
		}

		/// <summary>
		/// Serialize OpenXML document as binary array
		/// </summary>
		public byte[] ToArray()
		{
			_wordDocument.Close();
			return _documentStream.ToArray();
		}

		///<summary>
		/// Removes global readonly protection from OpenXML docx document
		///</summary>
		public void Unprotect()
		{
			SetProtectionAttribute("None");
		}

		///<summary>
		/// Sets global readonly protection for OpenXML docx document
		///</summary>
		public void Protect()
		{
			SetProtectionAttribute("readOnly");
		}

		private void SetProtectionAttribute(string value)
		{
			Settings settings = _wordDocument.MainDocumentPart.DocumentSettingsPart.Settings;
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

		///<summary>
		/// Sets value for specified <paramref name="key"/> into docx document inner key-value storage
		///</summary>
		///<param name="key">Key name</param>
		///<param name="value">Value to be set</param>
		public void SetCustomProperty(string key, string value)
		{
			if (_wordDocument.CustomFilePropertiesPart == null)
			{
				var addCustomFilePropertiesPart = _wordDocument.AddCustomFilePropertiesPart();
				addCustomFilePropertiesPart.Properties = new Properties();
			}

			var properties = _wordDocument.CustomFilePropertiesPart.Properties;

			var existsProperty = properties.OfType<CustomDocumentProperty>()
				.SingleOrDefault(x => x.Name.HasValue && x.Name.Value == key);
			if (existsProperty != null)
				existsProperty.Remove();

			Int32Value nextPropertyId = properties.Elements<CustomDocumentProperty>()
			                            	.Select(x => x.PropertyId.Value)
			                            	.Union(new[] {1})
			                            	.Max() + 1;

			properties.Append(new CustomDocumentProperty(new VTLPWSTR(value))
			                                                        	{
			                                                        		FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
			                                                        		PropertyId = nextPropertyId,
			                                                        		Name = key
			                                                        	});
		}

		///<summary>
		/// Gets value for specified <paramref name="key"/> from docx document inner key-value storage
		///</summary>
		///<param name="key">Key name</param>
		public string GetCustomProperty(string key)
		{
			if (_wordDocument.CustomFilePropertiesPart == null)
				return null;

			var openXmlElement = _wordDocument.CustomFilePropertiesPart.Properties
				.OfType<CustomDocumentProperty>()
				.FirstOrDefault(x => x.Name.HasValue && x.Name.Value == key);

			if (openXmlElement == null)
				return null;

			return openXmlElement.InnerText;
		}

		/// <summary>
		/// Placeholder fields created using <see cref="SdtElement"/> subclasses
		/// </summary>
		public FieldsCollection Fields
		{
			get { return _fields; }
		}

		///<summary>
		/// Appends paragraph with specified <paramref name="content"/> to the end of the document
		///</summary>
		///<param name="content">Text content</param>
		///<param name="visible">Specifies whether paragraph appended is visible or not</param>
		internal void AppendParagraph(string content, bool visible = true)
		{
			var paragraph = new Paragraph();
			var run = new Run(new Text(content));
			
			if (visible == false)
			{
				run.RunProperties = new RunProperties(new Vanish());
				paragraph.ParagraphProperties = new ParagraphProperties(new RunProperties(new Vanish()));
			}

			paragraph.AppendChild(run);

			_wordDocument.MainDocumentPart.Document.Body.Append(paragraph);
			_wordDocument.MainDocumentPart.Document.Save();
		}

		internal void InsertTagContent(string tagName, IEnumerable<TextBlock> content)
		{
			var paragraphContent = content.SelectMany(WrapText);
			var paragraph = new Paragraph(paragraphContent);

			InsertTagContent(tagName, paragraph);
		}

		internal void InsertTagContent(string tagName, OpenXmlElement content)
		{
			var tag = GetParagraph(GetCloseTagName(tagName));

			if (tag != null)
				tag.InsertBeforeSelf(content);

			_wordDocument.MainDocumentPart.Document.Save();
		}

	    internal WordprocessingDocument GetWordDocument()
		{
			return _wordDocument;
		}

		///<summary>
		/// Appends opening and closing tags pair to the end of docx document
		///</summary>
		///<param name="tagName">Tag name</param>
		public void AppendTag(string tagName)
		{
			AppendParagraph(GetOpenTagName(tagName), false);
			AppendParagraph(GetCloseTagName(tagName), false);
		}
	
		///<summary>
		/// Appends selfclosing tag to the end of docx document
		///</summary>
		///<param name="tagName">Tag name</param>
		///<param name="isVisible">Specifies whether new tag will be visible or not</param>
		public void AppendSelfclosingTag(string tagName, bool isVisible = true)
		{
			AppendParagraph(GetSingleTagName(tagName), isVisible);
		}

		private IEnumerable<OpenXmlElement> WrapText(TextBlock textBlock)
		{
			OpenXmlElement content;
			if (textBlock.Text == "\t") content = new TabChar();
			else content = new Text(textBlock.Text);

			var text = new Run(content);

			if (textBlock.Editable)
			{
				var id = _wordDocument.MainDocumentPart.Document.Descendants().OfType<PermStart>()
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