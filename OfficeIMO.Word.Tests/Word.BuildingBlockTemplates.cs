using System.IO.Compression;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Every built-in building block (cover pages, page numbers, tables of contents and the text watermark) and the
    /// default parts and styles of a new document must keep writing the same parts. The expected hashes in <c>Baselines/BuildingBlockTemplates.txt</c> were captured from
    /// the generated SDK builders before the templates moved to embedded markup. Set
    /// <c>OFFICEIMO_CAPTURE_BUILDING_BLOCK_BASELINE</c> to a file path to write the current output there instead, after
    /// an intended change to a template.
    /// </summary>
    public class WordBuildingBlockTemplates {
        private const string BaselineFile = "BuildingBlockTemplates.txt";

        [Fact]
        public void EveryTemplateWritesTheSamePartsAsTheCapturedBaseline() {
            var actual = new List<string>();
            foreach (WordCoverPageTemplate template in Enum.GetValues(typeof(WordCoverPageTemplate))) {
                actual.AddRange(Render($"CoverPage/{template}", document => document.AddCoverPage(template)));
            }
            foreach (WordPageNumberStyle style in Enum.GetValues(typeof(WordPageNumberStyle))) {
                actual.AddRange(Render($"PageNumber/Header/{style}", document => document.PageSetup.Header.AddPageNumber(style)));
                actual.AddRange(Render($"PageNumber/Footer/{style}", document => document.PageSetup.Footer.AddPageNumber(style)));
            }
            foreach (WordTableOfContentsStyle style in Enum.GetValues(typeof(WordTableOfContentsStyle))) {
                actual.AddRange(Render($"TableOfContent/{style}", document => document.AddTableOfContent(style)));
            }
            // The default parts and style definitions the templates sit in: a document created with another theme, the
            // theme part as generated, and a document without styles that gains them as its content asks for them.
            actual.AddRange(Render("Document/Classic", document => document.AddParagraph("Classic"), OfficeTheme.Classic));
            actual.AddRange(RenderDefaultTheme());
            actual.AddRange(RenderStyleless());
            actual.AddRange(Render("Watermark/Text", document => {
                document.AddParagraph("Section");
                _ = document.PageSetup.Header;
                document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Draft");
            }));

            string? capture = Environment.GetEnvironmentVariable("OFFICEIMO_CAPTURE_BUILDING_BLOCK_BASELINE");
            if (!string.IsNullOrEmpty(capture)) {
                File.WriteAllLines(capture, actual);
                return;
            }

            string[] expected = File.ReadAllLines(Path.Combine(FindTestProjectDirectory(), "Baselines", BaselineFile))
                .Where(line => line.Length > 0).ToArray();
            var missing = expected.Except(actual).ToList();
            var unexpected = actual.Except(expected).ToList();
            Assert.True(missing.Count == 0 && unexpected.Count == 0,
                "Building block output differs from the baseline." + Environment.NewLine
                + "Expected:" + Environment.NewLine + string.Join(Environment.NewLine, missing) + Environment.NewLine
                + "Actual:" + Environment.NewLine + string.Join(Environment.NewLine, unexpected));
        }

        // One line per part of the saved package (document properties excepted, they carry the save time): the
        // building block, the part name and the SHA-256 of its text with relationship ids numbered.
        private static IEnumerable<string> Render(string name, Action<WordDocument> build, OfficeTheme? theme = null) {
            byte[] package;
            using (var document = theme == null ? WordDocument.Create() : WordDocument.Create(new WordCreateOptions { Theme = theme })) {
                build(document);
                package = document.ToBytes();
            }
            return Hash(name, package);
        }

        private static IEnumerable<string> RenderDefaultTheme() {
            using var stream = new MemoryStream();
            using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document)) {
                MainDocumentPart main = package.AddMainDocumentPart();
                main.Document = new Document(new Body());
                WordDocument.GenerateThemePart1Content(main.AddNewPart<ThemePart>());
            }
            return Hash("DefaultPart/Theme", stream.ToArray());
        }

        private static IEnumerable<string> RenderStyleless() {
            using var input = new MemoryStream();
            using (var package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document)) {
                MainDocumentPart main = package.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(new Run(new Text("Styleless")))));
            }
            input.Position = 0;
            byte[] output;
            using (WordDocument document = WordDocument.Load(input)) {
                foreach (WordParagraphStyles style in Enum.GetValues(typeof(WordParagraphStyles))) {
                    if (style == WordParagraphStyles.Custom) continue;
                    document.AddParagraph(style.ToString()).SetStyle(style);
                }
                document.AddParagraph("Note").AddFootNote("Footnote");
                document.PageSetup.Header.AddPageNumber(WordPageNumberStyle.PlainNumber);
                document.PageSetup.Footer.AddPageNumber(WordPageNumberStyle.PlainNumber);
                output = document.ToBytes();
            }
            return Hash("Document/Styleless", output);
        }

        private static IEnumerable<string> Hash(string name, byte[] package) {
            using var stream = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var parts = new List<(string Name, string Text)>();
            foreach (ZipArchiveEntry entry in archive.Entries.OrderBy(entry => entry.FullName, StringComparer.Ordinal)) {
                if (entry.FullName.StartsWith("docProps/", StringComparison.Ordinal)) continue;
                using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
                parts.Add((entry.FullName, reader.ReadToEnd()));
            }

            // Relationship ids are generated at random; number them in package order so the hashes compare content.
            var ids = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var part in parts.Where(part => part.Name.EndsWith(".rels", StringComparison.Ordinal))) {
                foreach (Match match in Regex.Matches(part.Text, "\\bId=\"([^\"]+)\"")) {
                    string id = match.Groups[1].Value;
                    if (!ids.ContainsKey(id)) ids[id] = "rel" + ids.Count.ToString(CultureInfo.InvariantCulture);
                }
            }

            var lines = new List<string>();
            foreach (var part in parts) {
                string text = Regex.Replace(part.Text, "\"([^\"<>]+)\"", match =>
                    ids.TryGetValue(match.Groups[1].Value, out string? id) ? "\"" + id + "\"" : match.Value);
                // A new document's section gets a random revision id, and a table of contents a random placeholder name.
                text = Regex.Replace(text, "<w:sectPr w:rsidR=\"[0-9A-F]{8}\"", "<w:sectPr w:rsidR=\"00000000\"");
                text = Regex.Replace(text, "OfficeIMO_TocPlaceholder_[0-9a-f]{32}", "OfficeIMO_TocPlaceholder");
                string hash;
                using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", string.Empty);
                lines.Add($"{name} {part.Name} {hash}");
            }
            return lines;
        }

        private static string FindTestProjectDirectory() {
            string? directory = AppContext.BaseDirectory;
            while (directory != null) {
                string candidate = Path.Combine(directory, "OfficeIMO.Word.Tests");
                if (File.Exists(Path.Combine(candidate, "OfficeIMO.Word.Tests.csproj"))) return candidate;
                directory = Path.GetDirectoryName(directory);
            }
            throw new InvalidOperationException("The test project directory was not found above the test output.");
        }
    }
}
