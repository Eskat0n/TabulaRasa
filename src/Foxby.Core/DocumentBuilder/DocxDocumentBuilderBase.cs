using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Extensions;

namespace Foxby.Core.DocumentBuilder
{
    /// <summary>
	/// Contains common operations for working OpenXML tree
	/// </summary>
	public abstract class DocxDocumentBuilderBase
	{
		protected readonly WordprocessingDocument Document;
		private int _currentPermId;

		protected DocxDocumentBuilderBase(WordprocessingDocument document)
		{
			Document = document;
		}

	    protected OpenXmlElement CreatePermStart()
		{
			_currentPermId = Document.MainDocumentPart.Document
				.Descendants()
				.OfType<PermStart>()
				.Select(x => x.Id.Value)
				.Union(new[] {1})
				.Max();

			return new PermStart
			       	{
			       		Id = _currentPermId,
			       		EditorGroup = RangePermissionEditingGroupValues.Everyone
			       	};
		}

		protected OpenXmlElement CreatePermEnd()
		{
			return new PermEnd { Id = _currentPermId };
		}

        protected void SaveDocument()
		{
            foreach (var element in Document.MainDocumentPart.GetRootElements())
                element.Save();
		}

        protected static IEnumerable<Run> CreateTextContent(IEnumerable<string> content, RunProperties runProperties = null)
		{
			return content
				.SelectMany(x => x.Split(new[] {Environment.NewLine}, StringSplitOptions.None))
				.SelectMany((contentLine, index) => index == 0
				                                    	? new[]
				                                    	  	{
				                                    	  		new Run(new Text(contentLine) {Space = SpaceProcessingModeValues.Preserve})
				                                    	  			{
				                                    	  				RunProperties = CloneRunPropertiesIfNotNull(runProperties)
				                                    	  			}
				                                    	  	}
				                                    	: new[]
				                                    	  	{
				                                    	  		new Run(new Break()) {RunProperties = CloneRunPropertiesIfNotNull(runProperties)},
				                                    	  		new Run(new Text(contentLine) {Space = SpaceProcessingModeValues.Preserve})
				                                    	  			{
				                                    	  				RunProperties = CloneRunPropertiesIfNotNull(runProperties)
				                                    	  			}
				                                    	  	})
				.ToList();
		}

		protected static void ClearBetweenElements(OpenXmlElement start, OpenXmlElement end)
		{
            while (start.NextSibling() != end && start.NextSibling() != end)
				start.NextSibling().Remove();
		}

		private static RunProperties CloneRunPropertiesIfNotNull(RunProperties runProperties)
		{
			return runProperties == null ? null : runProperties.CloneElement();
		}
	}
}