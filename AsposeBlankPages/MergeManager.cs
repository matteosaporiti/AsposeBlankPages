using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Words;
using Document = Aspose.Words.Document;

namespace AsposeBlankPages
{
    class MergeManager
    {
        private static readonly Regex MergeRegex = new Regex(@"MERGE (""|“)((?!""|”|“).)*\.docx(""|”)");
        private static readonly string ProjDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

        internal static byte[] MergeWithAspose(byte[] doc, Dictionary<string, byte[]> attachments)
        {

            if (attachments == null || attachments.Count == 0)
                return doc;

            var licensePath = ProjDirectory + @"\lib\Aspose.Total.lic";
            AsposeLicenseManager.RegWordsLibrary(AsposeLicenseManager.LicenceType.Production, licensePath, true);

            using (var outputStream = new MemoryStream())
            {
                using (var docStream = new MemoryStream(doc))
                {
                    var currentDoc = new Document(docStream);
                    string text = currentDoc.Cast<Section>().Aggregate("", (current, section) => current + section.GetText());

                    var matches = MergeRegex.Matches(text);
                    var matchedAttachments = GetMatchedAttachments(attachments, matches);

                    var outputDoc = MergeAttachmentsInOutputDoc(currentDoc, matchedAttachments);
                    outputDoc.UpdatePageLayout();
                    outputDoc.Save(outputStream, SaveFormat.Docx);
                    outputDoc.Save(ProjDirectory + @"\output\Test_generated.docx", Aspose.Words.SaveFormat.Docx);
                    outputDoc.Save(ProjDirectory + @"\output\Test_generated.pdf", Aspose.Words.SaveFormat.Pdf);
                }
                return outputStream.ToArray();
            }
        }

        private static Document MergeAttachmentsInOutputDoc(Document asposeDoc, Dictionary<string, byte[]> matchedAttachments)
        {
            var outputDoc = (Document)asposeDoc.Clone(false);

            int start = 0;
            int currentIndex = 0;
            var mainSection = asposeDoc.FirstSection;
            var width = mainSection.PageSetup.PageWidth;
            var heigth = mainSection.PageSetup.PageHeight;
            var mainSections = new Queue<int>();
            foreach (Node par in mainSection.Body)
            {
                string found = matchedAttachments
                    .Select(x => x.Key)
                    .FirstOrDefault(x => par.GetText().Contains(x));
                if (found != null)
                {
                    InsertSlicedSectionIntoOutput(asposeDoc, start, currentIndex, outputDoc, mainSections);
                    InsertAttachmentDocIntoOutput(matchedAttachments[found], outputDoc, width, heigth);
                    mainSections.Enqueue(0);
                    start = currentIndex + 1;
                }
                currentIndex++;
            }
            if (currentIndex > start)
            {
                InsertSlicedSectionIntoOutput(asposeDoc, start, currentIndex, outputDoc, mainSections);
            }
            //To restart from 1 each time
            foreach (Section section in outputDoc)
            {
                ////////////////
                /// 
                /// 
                /// Comment the following line and the bug is gone
                /// 
                /// 
                /// 
                ////////////////
                section.PageSetup.RestartPageNumbering = true; 
            }
            return outputDoc;
        }

        private static void InsertAttachmentDocIntoOutput(byte[] attachment, Document outputDoc, double w, double h)
        {
            using (var mergedStream = new MemoryStream(attachment))
            {
                var attachmentDocument = new Document(mergedStream);

                attachmentDocument.FirstSection.HeadersFooters.LinkToPrevious(false);
                attachmentDocument.FirstSection.PageSetup.PageWidth = w;
                attachmentDocument.FirstSection.PageSetup.PageHeight = h;
                var sectionFromAttachment = outputDoc.ImportNode(attachmentDocument.FirstSection, true, ImportFormatMode.KeepSourceFormatting);

                outputDoc.AppendChild(sectionFromAttachment);
            }
        }

        private static void InsertSlicedSectionIntoOutput(Document asposeDoc, int start, int currentIndex, Document outputDoc, Queue<int> mainSections)
        {
            var slicedSection = GetSliceOfSection((Section)asposeDoc.FirstSection.Clone(true), start, currentIndex);
            if (IsNullOrWhiteSpaceOrLineBreak(slicedSection.Body.GetText())) return; // don't add empty sections

            slicedSection.HeadersFooters.LinkToPrevious(false);
            var slicedSectionToInsert = outputDoc.ImportNode(slicedSection, true, ImportFormatMode.KeepSourceFormatting);

            outputDoc.AppendChild(slicedSectionToInsert);
        }

        private static Dictionary<string, byte[]> GetMatchedAttachments(Dictionary<string, byte[]> attachments, MatchCollection matches)
        {
            var matchedAttachments = new Dictionary<string, byte[]>();

            foreach (Match match in matches)
            {
                string filename = null;
                foreach (var segment in match.Value.Split(new[] { '“', '"', '”' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (segment.IndexOf(".docx", StringComparison.OrdinalIgnoreCase) != -1) //case insensitive
                    {
                        filename = segment;
                    }
                }
                if (filename != null && attachments.Keys.Contains(filename))
                {
                    matchedAttachments.Add(match.Value, attachments[filename]);
                }
            }
            return matchedAttachments;
        }

        private static Section GetSliceOfSection(Section clonedSection, int start, int currentIndex)
        {
            int i = 0;
            foreach (Node node in clonedSection.Body)
            {
                if (i < start || i >= currentIndex)
                {
                    node.Remove();
                }
                i++;
            }
            return clonedSection;
        }

        private static bool IsNullOrWhiteSpaceOrLineBreak(string text)
        {
            var noBreaks = Regex.Replace(text, @"\r\n?|\n", "");
            return string.IsNullOrWhiteSpace(noBreaks);
        }

        
    }
}
