using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordFieldUpdateReportTests {
        private static readonly XNamespace FieldFixtureWord = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace FieldFixtureMc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private static readonly XNamespace FieldFixtureVendor = "urn:officeimo:field-fixture:vendor";

        [Theory]
        [InlineData(false, false, "carrier", "2")]
        [InlineData(false, true, "carrier", "2")]
        [InlineData(true, false, "carrier", "2999")]
        [InlineData(true, false, "extLst", "2999")]
        public void Test_FieldCompatibility_FormulaUsesEffectiveCellContent(bool processContent, bool nestedTable, string wrapper, string expected) {
            byte[] bytes;
            using (WordDocument source = WordDocument.Create()) {
                WordTable table = source.AddTable(2, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "2";
                table.Rows[1].Cells[0].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale"));
                bytes = source.ToBytes();
            }

            bytes = EditFieldFixturePart(bytes, name => name == "word/document.xml", root => {
                XNamespace w = FieldFixtureWord;
                SetFieldFixtureIgnorableNamespace(root);
                if (processContent) root.SetAttributeValue(FieldFixtureMc + "ProcessContent", "v:" + wrapper);
                XElement cell = root.Descendants(w + "tc").First();
                if (nestedTable) {
                    XElement table = new XElement(w + "tbl", new XElement(w + "tr",
                        new XElement(w + "tc", new XElement(w + "p"))));
                    cell.Add(new XElement(FieldFixtureVendor + wrapper, table));
                } else {
                    cell.Element(w + "p")!.Add(new XElement(FieldFixtureVendor + wrapper,
                        new XElement(w + "r", new XElement(w + "t", "999"))));
                }
            });

            foreach (bool sdk in new[] { false, true }) {
                using WordDocument document = LoadFieldCompatibilityFixture(bytes, sdk);
                WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);
                Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
                Assert.Equal(expected, result.ResultText);
                Assert.Equal(expected, Assert.Single(document.InspectFields()).ResultText);
                AssertFieldCompatibilityBackend(document, sdk);
            }
        }

        [Theory]
        [InlineData(false, "1234", WordFieldUpdateStatus.Updated)]
        [InlineData(true, "stale", WordFieldUpdateStatus.Skipped)]
        public void Test_FieldCompatibility_CustomPropertyMatchesDirectSdkValues(bool wrapValue, string expected, WordFieldUpdateStatus status) {
            byte[] bytes;
            using (WordDocument source = WordDocument.Create()) {
                source.CustomProperties["Amount"] = new WordCustomProperty(1234);
                source.AddParagraph()._paragraph.Append(BuildSimpleField(" DOCPROPERTY Amount ", "stale"));
                bytes = source.ToBytes();
            }

            bytes = EditFieldFixturePart(bytes, name => name == "docProps/custom.xml", root => {
                SetFieldFixtureIgnorableNamespace(root);
                XElement property = Assert.Single(root.Elements(root.Name.Namespace + "property"));
                XElement scalar = Assert.Single(property.Elements());
                property.AddFirst(new XElement(FieldFixtureVendor + "extLst", new XElement(scalar.Name, "999")));
                if (wrapValue) {
                    root.SetAttributeValue(FieldFixtureMc + "ProcessContent", "v:carrier");
                    scalar.ReplaceWith(new XElement(FieldFixtureVendor + "carrier", new XElement(scalar)));
                }
            });

            foreach (bool sdk in new[] { true, false }) {
                using WordDocument document = LoadFieldCompatibilityFixture(bytes, sdk);
                WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);
                Assert.Equal(status, result.Status);
                Assert.Equal(expected, Assert.Single(document.InspectFields()).ResultText);
                AssertFieldCompatibilityBackend(document, sdk);
            }
        }

        [Theory]
        [InlineData(" \t\n  ", "stale", WordFieldUpdateStatus.Skipped)]
        [InlineData("  Quarterly review  ", "  Quarterly review  ", WordFieldUpdateStatus.Updated)]
        public void Test_FieldCompatibility_TitleMatchesPackageWhitespaceRules(string title, string expected, WordFieldUpdateStatus status) {
            byte[] bytes = CreateCoreFieldFixture(false, title);
            foreach (bool sdk in new[] { true, false }) {
                using WordDocument document = LoadFieldCompatibilityFixture(bytes, sdk);
                WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);
                Assert.Equal(status, result.Status);
                Assert.Equal(expected, Assert.Single(document.InspectFields()).ResultText);
                AssertFieldCompatibilityBackend(document, sdk);
            }
        }

        [Theory]
        [InlineData("2024-02-03")]
        [InlineData("2024-02-03 04:05:06")]
        [InlineData("2024-02-03T04:05:06.1")]
        [InlineData("2024-02-03T04:05:06.12345678Z")]
        public void Test_FieldCompatibility_InvalidCreatedTimestampIsRejected(string timestamp) {
            byte[] bytes = CreateCoreFieldFixture(true, timestamp);
            foreach (bool sdk in new[] { true, false }) {
                using WordDocument document = LoadFieldCompatibilityFixture(bytes, sdk);
                Assert.Throws<XmlException>(() => document.UpdateFieldsAndGetReport());
                Assert.Equal("stale", Assert.Single(document.InspectFields()).ResultText);
                AssertFieldCompatibilityBackend(document, sdk);
            }
        }

        [Theory]
        [InlineData("2024-02-03T04:05:06")]
        [InlineData("2024-02-03T04:05:06Z")]
        [InlineData("2024-02-03T04:05:06.1+05:30")]
        public void Test_FieldCompatibility_ValidCreatedTimestampMatchesSdk(string timestamp) {
            byte[] bytes = CreateCoreFieldFixture(true, timestamp);
            using WordDocument sdk = LoadFieldCompatibilityFixture(bytes, true);
            using WordDocument owned = LoadFieldCompatibilityFixture(bytes, false);

            WordFieldUpdateResult expected = Assert.Single(sdk.UpdateFieldsAndGetReport().Results);
            WordFieldUpdateResult actual = Assert.Single(owned.UpdateFieldsAndGetReport().Results);

            Assert.Equal(WordFieldUpdateStatus.Updated, expected.Status);
            Assert.NotEqual("stale", expected.ResultText);
            Assert.Equal(expected.Status, actual.Status);
            Assert.Equal(expected.ResultText, actual.ResultText);
            Assert.Equal(expected.ResultText, Assert.Single(owned.InspectFields()).ResultText);
            AssertFieldCompatibilityBackend(owned, false);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_FieldCompatibility_TitleRejectsAnInvalidSiblingTimestamp(bool cdata) {
            byte[] bytes = CreateCoreFieldFixture(false, "Valid title", root => {
                XNamespace terms = "http://purl.org/dc/terms/";
                XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
                root.SetAttributeValue(XNamespace.Xmlns + "dcterms", terms.NamespaceName);
                XText content = cdata ? new XCData("2024-02-03T04:05:06Z") : new XText("2024-02-03");
                root.Add(new XElement(terms + "created", new XAttribute(xsi + "type", "dcterms:W3CDTF"), content));
            });

            foreach (bool sdk in new[] { true, false }) {
                using WordDocument document = LoadFieldCompatibilityFixture(bytes, sdk);
                Assert.Throws<XmlException>(() => document.UpdateFieldsAndGetReport());
                Assert.Equal("stale", Assert.Single(document.InspectFields()).ResultText);
                AssertFieldCompatibilityBackend(document, sdk);
            }
        }

        private static byte[] CreateCoreFieldFixture(bool date, string value, Action<XElement>? edit = null) {
            XNamespace ns = date ? "http://purl.org/dc/terms/" : "http://purl.org/dc/elements/1.1/";
            string property = date ? "created" : "title";
            byte[] bytes;
            using (WordDocument source = WordDocument.Create()) {
                if (date) source.Properties.Created = new DateTime(2024, 2, 3, 4, 5, 6, DateTimeKind.Utc);
                else source.Properties.Title = "Seed title";
                string instruction = date ? " CREATEDATE \\@ \"yyyy-MM-dd HH:mm:ss\" " : " TITLE ";
                source.AddParagraph()._paragraph.Append(BuildSimpleField(instruction, "stale"));
                bytes = source.ToBytes();
            }

            return EditFieldFixturePart(bytes, name => name == "docProps/core.xml" ||
                name.EndsWith(".psmdcp", StringComparison.Ordinal), root => {
                    root.Element(ns + property)!.Value = value;
                    edit?.Invoke(root);
                });
        }

        private static byte[] EditFieldFixturePart(byte[] bytes, Func<string, bool> part, Action<XElement> edit) {
            using var output = new MemoryStream();
            output.Write(bytes, 0, bytes.Length);
            using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry entry = Assert.Single(archive.Entries, item => part(item.FullName));
                XDocument xml;
                using (Stream input = entry.Open()) xml = XDocument.Load(input);
                edit(xml.Root!);
                string name = entry.FullName;
                entry.Delete();
                using Stream replacement = archive.CreateEntry(name).Open();
                xml.Save(replacement, SaveOptions.DisableFormatting);
            }
            return output.ToArray();
        }

        private static void SetFieldFixtureIgnorableNamespace(XElement root) {
            root.SetAttributeValue(XNamespace.Xmlns + "v", FieldFixtureVendor.NamespaceName);
            XName ignorable = FieldFixtureMc + "Ignorable";
            string? existing = (string?)root.Attribute(ignorable);
            root.SetAttributeValue(ignorable, existing == null ? "v" : existing + " v");
        }

        private static WordDocument LoadFieldCompatibilityFixture(byte[] bytes, bool sdk) {
            WordDocument document = WordDocument.Load(new MemoryStream(bytes));
            if (sdk) _ = document.OpenXmlDocument;
            AssertFieldCompatibilityBackend(document, sdk);
            return document;
        }

        private static void AssertFieldCompatibilityBackend(WordDocument document, bool sdk) {
#if NET
            Assert.Equal(!sdk, document.IsSdkDeferred);
#endif
        }
    }
}
