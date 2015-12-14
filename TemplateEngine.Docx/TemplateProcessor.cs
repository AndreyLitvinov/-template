using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using TemplateEngine.Docx.Processors;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;


namespace TemplateEngine.Docx
{
	public class TemplateProcessor : IDisposable
	{
		public readonly XDocument Document;
		public readonly XDocument NumberingPart;
		public readonly XDocument StylesPart;
		private readonly WordprocessingDocument _wordDocument;
		private bool _isNeedToRemoveContentControls;
		private bool _isNeedToNoticeAboutErrors;

		public TemplateProcessor(string fileName)
		{
			_wordDocument = WordprocessingDocument.Open(fileName, true);
			_isNeedToNoticeAboutErrors = true;

			Document = LoadPart(_wordDocument.MainDocumentPart);
			NumberingPart = LoadPart(_wordDocument.MainDocumentPart.NumberingDefinitionsPart);
			StylesPart = LoadPart(_wordDocument.MainDocumentPart.StyleDefinitionsPart);

		}

		public TemplateProcessor(Stream stream)
		{
			_wordDocument = WordprocessingDocument.Open(stream, true);
			_isNeedToNoticeAboutErrors = true;

			Document = LoadPart(_wordDocument.MainDocumentPart);
			NumberingPart = LoadPart(_wordDocument.MainDocumentPart.NumberingDefinitionsPart);
			StylesPart = LoadPart(_wordDocument.MainDocumentPart.StyleDefinitionsPart);

		}

		public TemplateProcessor(XDocument templateSource, XDocument stylesPart = null, XDocument numberingPart = null)
		{
			_isNeedToNoticeAboutErrors = true;

			Document = templateSource;
			StylesPart = stylesPart;
			NumberingPart = numberingPart;
		}

		private XDocument LoadPart(OpenXmlPart source)
		{
			if (source == null) return null;

			var part = source.Annotation<XDocument>();
			if (part != null) return part;

			using (var str = source.GetStream())
			using (var streamReader = new StreamReader(str))
			using (var xr = XmlReader.Create(streamReader))
				part = XDocument.Load(xr);

			return part;
		}
		public TemplateProcessor SetRemoveContentControls(bool isNeedToRemove)
		{
			_isNeedToRemoveContentControls = isNeedToRemove;
			return this;
		}
		public TemplateProcessor SetNoticeAboutErrors(bool isNeedToNotice)
		{
			_isNeedToNoticeAboutErrors = isNeedToNotice;
			return this;
		}

		public void SaveChanges()
		{
			if (Document == null) return;
			// Serialize the XDocument object back to the package.
			using (var xw = XmlWriter.Create(_wordDocument.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
			{
				Document.Save(xw);
			}

			if (NumberingPart != null)
			{
				// Serialize the XDocument object back to the package.
				using (var xw = XmlWriter.Create(_wordDocument.MainDocumentPart.NumberingDefinitionsPart.GetStream(FileMode.Create,
							FileAccess.Write)))
				{
					NumberingPart.Save(xw);
				}
			}
			_wordDocument.Close();
		}

		public void SaveChanges(Func<string, bool> AliasName)
		{
			if (Document == null) return;
			// Serialize the XDocument object back to the package.
			using (var xw = XmlWriter.Create(_wordDocument.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
			{
				Document.Save(xw);
			}
			if (NumberingPart != null)
			{
				// Serialize the XDocument object back to the package.
				using (var xw = XmlWriter.Create(_wordDocument.MainDocumentPart.NumberingDefinitionsPart.GetStream(FileMode.Create,
							FileAccess.Write)))
				{
					NumberingPart.Save(xw);
				}
			}

			MainDocumentPart mainDocumentPart = _wordDocument.MainDocumentPart;
			List<SdtBlock> sdtList = mainDocumentPart.Document.Descendants<SdtBlock>().ToList();
			List<SdtElement> newListSdt = findeStd(mainDocumentPart.Document);
			SdtElement sdtA = null;
			foreach (SdtElement sdt in newListSdt)
			{
				sdtA = sdt;
				if (AliasName(sdtA.SdtProperties.GetFirstChild<SdtAlias>().Val.Value))
				{
					Paragraph parentP = sdt.Parent as Paragraph;
					if (sdt.Parent != null)
					{
						sdtA.Remove();
					}
					if (parentP != null)
					{
						parentP.Remove();
					}

				}
			}
			_wordDocument.Close();
		}

		private List<SdtElement> findeStd(OpenXmlElement mainDocumentPart)
		{
			List<SdtElement> result = new List<SdtElement>();
			var tmpList = mainDocumentPart.Descendants<OpenXmlElement>().ToList();
			if (tmpList.Count != 0)
			{
				foreach (OpenXmlElement item in tmpList)
				{
					SdtElement sdt = item as SdtElement;
					if (sdt != null)
					{
						result.Add(sdt);
					}
					else
					{
						result.AddRange(findeStd(item));
					}
				}
			}
			return result;
		}

		public TemplateProcessor FillContent(Content content)
		{
			var processResult =
				new ContentProcessor(
					new ProcessContext(Document, NumberingPart, StylesPart))
					.SetRemoveContentControls(_isNeedToRemoveContentControls)
					.FillContent(Document.Root.Element(W.body), content);

			if (_isNeedToNoticeAboutErrors)
				AddErrors(processResult.Errors);

			return this;
		}

		public static bool getTest() { return true; }

		// Add any errors as red text on yellow at the beginning of the document.
		private void AddErrors(IList<string> errors)
		{
			if (errors.Any())
				Document.Root
					.Element(W.body)
					.AddFirst(errors.Select(s =>
						new XElement(W.p,
							new XElement(W.r,
								new XElement(W.rPr,
									new XElement(W.color,
										new XAttribute(W.val, "red")),
									new XElement(W.sz,
										new XAttribute(W.val, "28")),
									new XElement(W.szCs,
										new XAttribute(W.val, "28")),
									new XElement(W.highlight,
										new XAttribute(W.val, "yellow"))),
								new XElement(W.t, s)))));
		}

		public void Dispose()
		{
			if (_wordDocument == null) return;

			_wordDocument.Dispose();
		}
	}
}
