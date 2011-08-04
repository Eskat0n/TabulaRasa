using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Foxby.Core.DocumentBuilder.Extensions;

namespace Foxby.Core.DocumentBuilder
{
	public abstract class DocxDocumentBuilderBase
	{
		protected readonly WordprocessingDocument Document;
		private int currentPermId;

		protected DocxDocumentBuilderBase(WordprocessingDocument document)
		{
			Document = document;
		}

		protected OpenXmlElement CreatePermStart()
		{
			currentPermId = Document.MainDocumentPart.Document
				.Descendants()
				.OfType<PermStart>()
				.Select(x => x.Id.Value)
				.Union(new[] {1})
				.Max();

			return new PermStart
			       	{
			       		Id = currentPermId,
			       		EditorGroup = RangePermissionEditingGroupValues.Everyone
			       	};
		}

		protected OpenXmlElement CreatePermEnd()
		{
			return new PermEnd { Id = currentPermId };
		}

		protected void SaveDocument()
		{
			Document.MainDocumentPart.Document.Save();
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