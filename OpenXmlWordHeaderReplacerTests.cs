using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WordImageReplace.Tests
{
    [TestClass]
    public class OpenXmlWordHeaderReplacerTests
    {
        private const string TestDocPath = "test.docx";
        private const string DefaultHeaderImage = "default.png";
        private const string EvenHeaderImage = "even.png";

        [TestInitialize]
        public void Setup()
        {
            // テスト用の .docx と画像ファイルを準備（必要なら生成）
            File.Copy("minimal_template.docx", TestDocPath, true);
            File.Copy("sample_default.png", DefaultHeaderImage, true);
            File.Copy("sample_even.png", EvenHeaderImage, true);
        }

        [TestCleanup]
        public void Cleanup()
        {
            File.Delete(TestDocPath);
            File.Delete(DefaultHeaderImage);
            File.Delete(EvenHeaderImage);
        }

        [TestMethod]
        public void ReplaceHeaderImages_UpdatesDefaultAndEvenHeaders()
        {
            OpenXmlWordHeaderReplacer.ReplaceHeaderImages(
                TestDocPath,
                "output.docx",
                DefaultHeaderImage,
                EvenHeaderImage,
                changeFront: true,
                changeBack: true);

            using (var doc = WordprocessingDocument.Open("output.docx", false))
            {
                var mainPart = doc.MainDocumentPart;
                var sectionProps = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().ToList();

                bool defaultHeaderUpdated = false;
                bool evenHeaderUpdated = false;

                foreach (var sec in sectionProps)
                {
                    foreach (var headerRef in sec.Elements<DocumentFormat.OpenXml.Wordprocessing.HeaderReference>())
                    {
                        var headerPart = (HeaderPart)mainPart.GetPartById(headerRef.Id.Value);
                        var blip = headerPart.Header.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                        if (headerRef.Type?.Value == DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default)
                        {
                            defaultHeaderUpdated = blip != null;
                        }
                        else if (headerRef.Type?.Value == DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Even)
                        {
                            evenHeaderUpdated = blip != null;
                        }
                    }
                }

                Assert.IsTrue(defaultHeaderUpdated, "Default header image was not updated.");
                Assert.IsTrue(evenHeaderUpdated, "Even header image was not updated.");
            }
        }
    }
}