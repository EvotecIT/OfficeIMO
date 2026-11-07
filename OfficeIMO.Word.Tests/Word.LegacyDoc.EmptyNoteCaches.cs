using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using System.Text;
using Xunit;
using M = DocumentFormat.OpenXml.Math;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyDoc_EmptyFieldCaches_ReadableCommentsAndBodyTextboxesReportOmittedFields(bool textBox, bool separator) {
        foreach (string instruction in new[] { "DATE", "FILENAME", "DOCPROPERTY Title", "EQ \\f(1,2)", "HYPERLINK \"https://example.com\"", "PAGE", "NUMPAGES", "SECTIONPAGES" }) {
            bool acceptsEmpty = instruction == "SECTIONPAGES";
            foreach (string cache in new[] { string.Empty, "CACHE" }) {
                string field = "\u0013 " + instruction + " " + (separator ? "\u0014" + cache : cache) + "\u0015";
                if (!separator && cache.Length != 0) continue;
                byte[] bytes = CreateReadableStoryFieldCarrier(textBox, field, separator);
                using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
                string visible = textBox ? Assert.Single(loaded.TextBoxes).Content!.InnerText : Assert.Single(loaded.Comments).Text;
                string expectedCache = cache.Length != 0 ? cache : acceptsEmpty ? "1" : string.Empty;
                Assert.Equal("Before " + expectedCache + " After", visible);
                var findings = loaded.LegacyDocUnsupportedFeatures;
                using var output = new MemoryStream();
                if (!acceptsEmpty && cache.Length == 0) {
                    var finding = Assert.Single(findings);
                    Assert.Equal("doc-field-empty-result-unprojected", finding.Code);
                    Assert.Equal(textBox ? "Field:TextBox:EmptyResult" : "Field:Comment:EmptyResult", finding.DetailCode);
                    Assert.Contains(instruction, finding.Description);
                    Assert.Throws<NotSupportedException>(() => loaded.Save(output));
                    Assert.Equal(0, output.Length);
                    loaded.Save(output, new WordSaveOptions { AllowLoss = true });
                } else {
                    Assert.Empty(findings);
                    loaded.Save(output);
                }
                using var package = new System.IO.Compression.ZipArchive(new MemoryStream(output.ToArray()));
                string part = textBox ? "word/document.xml" : "word/comments.xml";
                using Stream xml = package.GetEntry(part)!.Open();
                var root = System.Xml.Linq.XDocument.Load(xml);
                System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var fields = root.Descendants(w + "fldSimple").ToArray();
                if (!acceptsEmpty && cache.Length == 0) {
                    Assert.Empty(fields);
                    Assert.Empty(root.Descendants(w + "hyperlink"));
                } else if (instruction.StartsWith("HYPERLINK ", StringComparison.Ordinal)) {
                    Assert.Empty(fields);
                    var hyperlink = Assert.Single(root.Descendants(w + "hyperlink"));
                    Assert.Equal(cache, hyperlink.Value);
                    System.Xml.Linq.XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    string relationshipId = (string)hyperlink.Attribute(r + "id")!;
                    string relationshipsPart = textBox ? "word/_rels/document.xml.rels" : "word/_rels/comments.xml.rels";
                    using Stream relationshipsXml = package.GetEntry(relationshipsPart)!.Open();
                    var relationships = System.Xml.Linq.XDocument.Load(relationshipsXml);
                    var relationship = Assert.Single(relationships.Root!.Elements(), element => (string?)element.Attribute("Id") == relationshipId);
                    Assert.Equal("External", (string?)relationship.Attribute("TargetMode"));
                    Assert.Equal(new Uri("https://example.com"), new Uri((string)relationship.Attribute("Target")!));
                    if (!textBox) {
                        using WordDocument reopened = WordDocument.Load(new MemoryStream(output.ToArray()));
                        WordParagraph linked = Assert.Single(Assert.Single(reopened.Comments).Paragraphs, paragraph => paragraph.Hyperlink != null);
                        Assert.Equal(new Uri("https://example.com"), linked.Hyperlink!.Uri);
                        linked.Hyperlink.Uri = new Uri("https://example.org/edited");
                        using var edited = new MemoryStream();
                        reopened.Save(edited);
                        using WordDocument editedReopened = WordDocument.Load(new MemoryStream(edited.ToArray()));
                        WordParagraph editedLink = Assert.Single(Assert.Single(editedReopened.Comments).Paragraphs, paragraph => paragraph.Hyperlink != null);
                        Assert.Equal(new Uri("https://example.org/edited"), editedLink.Hyperlink!.Uri);
                    }
                } else {
                    var projected = Assert.Single(fields);
                    Assert.Equal(instruction, ((string?)projected.Attribute(w + "instr"))?.Trim());
                    Assert.Equal(expectedCache, projected.Value);
                }
            }
        }

        byte[] emptyOnly = CreateReadableStoryFieldCarrier(textBox, "\u0013 DATE " + (separator ? "\u0014" : "") + "\u0015", separator, surroundingText: false);
        using (WordDocument loaded = WordDocument.Load(new MemoryStream(emptyOnly))) {
            string expected = textBox ? "doc-textbox-stories-present" : "doc-field-empty-result-unprojected";
            Assert.Equal(expected, Assert.Single(loaded.LegacyDocUnsupportedFeatures).Code);
            using var output = new MemoryStream();
            Assert.Throws<NotSupportedException>(() => loaded.Save(output));
        }
        byte[] unprojectedStory = CreateReadableStoryFieldCarrier(textBox, "\u0013 DATE " + (separator ? "\u0014" : "") + "\u0015", separator);
        Assert.True(OfficeCompoundFileReader.TryRead(unprojectedStory, out OfficeCompoundFile? generic, out _));
        byte[] word = generic!.Streams["WordDocument"];
        if (textBox) {
            Buffer.BlockCopy(word, 0x64, word, 0x68, 4);
            Array.Clear(word, 0x64, 4);
        } else Array.Clear(word, 0xBE, 4); // Unreadable comment reference PLC.
        byte[] guarded = OfficeCompoundFileWriter.Write(generic.Streams.Select(stream => new OfficeCompoundStream(stream.Key, stream.Value)).ToArray());
        using (WordDocument loaded = WordDocument.Load(new MemoryStream(guarded))) {
            Assert.Equal(textBox ? "doc-header-textbox-stories-present" : "doc-comment-stories-present", Assert.Single(loaded.LegacyDocUnsupportedFeatures).Code);
            Assert.DoesNotContain(loaded.LegacyDocImportDiagnostics, diagnostic => diagnostic.Code == "doc-field-empty-result-unprojected");
            using var output = new MemoryStream();
            Assert.Throws<NotSupportedException>(() => loaded.Save(output));
        }
    }

    private static byte[] CreateReadableStoryFieldCarrier(bool textBox, string field, bool separator, bool surroundingText = true) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Body");
        int prefixLength = surroundingText ? 7 : 0;
        string placeholder = (surroundingText ? "Before " : "") + new string('X', field.Length) + (surroundingText ? " After" : "");
        if (textBox) source.AddParagraph(placeholder);
        else source.Paragraphs[0].AddComment("Author", "AR", placeholder);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            source.Save(path);
            Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(path), out OfficeCompoundFile? compound, out _));
            byte[] main = compound!.Streams["WordDocument"].ToArray();
            byte[] table = compound.Streams["1Table"].ToArray();
            string text = Encoding.ASCII.GetString(main);
            int offset = text.IndexOf(new string('X', field.Length), StringComparison.Ordinal);
            Assert.True(offset >= 0);
            Encoding.ASCII.GetBytes(field).CopyTo(main, offset);
            int storyCount = BitConverter.ToInt32(main, textBox ? 0x4C : 0x5C);
            if (textBox) {
                Buffer.BlockCopy(BitConverter.GetBytes(5), 0, main, 0x4C, 4);
                storyCount -= 5;
                Buffer.BlockCopy(BitConverter.GetBytes(storyCount), 0, main, 0x64, 4);
            }
            int end = prefixLength + field.Length - 1;
            int split = field.IndexOf('\u0014');
            int[] cps = separator ? new[] { prefixLength, prefixLength + split, end } : new[] { prefixLength, end };
            int start = table.Length;
            int length = 4 + cps.Length * 6;
            Array.Resize(ref table, start + length);
            for (int index = 0; index < cps.Length; index++) Buffer.BlockCopy(BitConverter.GetBytes(cps[index]), 0, table, start + index * 4, 4);
            Buffer.BlockCopy(BitConverter.GetBytes(storyCount), 0, table, start + cps.Length * 4, 4);
            int properties = start + (cps.Length + 1) * 4;
            for (int index = 0; index < cps.Length; index++) {
                table[properties + index * 2] = (byte)(index == 0 ? 0x13 : index == cps.Length - 1 ? 0x15 : 0x14);
                table[properties + index * 2 + 1] = (byte)(index == cps.Length - 1 && separator ? 0x80 : 0);
            }
            int directory = textBox ? 0x262 : 0x132;
            Buffer.BlockCopy(BitConverter.GetBytes(start), 0, main, directory, 4);
            Buffer.BlockCopy(BitConverter.GetBytes(length), 0, main, directory + 4, 4);
            return OfficeCompoundFileWriter.Write(compound.Streams.Select(stream => new OfficeCompoundStream(stream.Key,
                stream.Key == "WordDocument" ? main : stream.Key == "1Table" ? table : stream.Value)).ToArray());
        } finally { DeleteIfExists(path); }
    }

    [Theory]
    [InlineData("FILENAME")]
    [InlineData("STYLEREF \"Heading 1\"")]
    [InlineData("DATE")]
    [InlineData("DOCPROPERTY Title")]
    [InlineData("COUNTS")]
    [InlineData("OMML")]
    public void LegacyDoc_EmptyFieldCaches_KeepNativeResultsAndReportUnprojectedImportsAcrossStories(string profile) {
        using WordDocument source = WordDocument.Create();
        WordFootNote footnote = source.AddParagraph("Footnote anchor").AddFootNote("Note ");
        WordEndNote endnote = source.AddParagraph("Endnote anchor").AddEndNote("Note ");
        var main = source._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement[] notes = {
            main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Single(note => note.Id?.Value == footnote.ReferenceId),
            main.EndnotesPart!.Endnotes!.Elements<W.Endnote>().Single(note => note.Id?.Value == endnote.ReferenceId)
        };
        var paragraphs = new List<W.Paragraph> {
            source.AddParagraph("Body fields ")._paragraph,
            source.Sections[0].PageSetup.Header.AddParagraph("Header fields ")._paragraph,
            source.Sections[0].PageSetup.Footer.AddParagraph("Footer fields ")._paragraph
        };
        paragraphs.AddRange(notes.Select(note => note.Elements<W.Paragraph>().First()));
        string[] instructions = profile == "COUNTS" ? new[] { "PAGE", "NUMPAGES", "SECTIONPAGES" } : new[] { profile };
        foreach (W.Paragraph paragraph in paragraphs) {
            if (profile == "OMML") {
                paragraph.Append(new M.OfficeMath(new M.Run(new M.Text(string.Empty))),
                    new M.Paragraph(new M.OfficeMath(new M.Run(new M.Text(string.Empty)))));
                continue;
            }
            foreach (string instruction in instructions) {
                paragraph.Append(new W.SimpleField { Instruction = " " + instruction + " ", FieldLock = profile != "COUNTS" },
                    new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
                    new W.Run(new W.FieldCode(" " + instruction + " ")),
                    new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
                    new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
            }
        }

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            source.Save(path);
            byte[] encoded = File.ReadAllBytes(path);
            Assert.True(OfficeCompoundFileReader.TryRead(encoded, out OfficeCompoundFile? compound, out _));
            byte[] stream = compound!.Streams["WordDocument"];
            int start = BitConverter.ToInt32(stream, 0x18);
            int end = BitConverter.ToInt32(stream, 0x1C);
            // The fixture is ASCII-only, so the native writer uses compressed text.
            string text = Encoding.ASCII.GetString(stream, start, end - start);
            var fields = new List<(string Instruction, string Result)>();
            for (int begin = text.IndexOf('\u0013'); begin >= 0;) {
                int separator = text.IndexOf('\u0014', begin + 1);
                int finish = text.IndexOf('\u0015', separator + 1);
                Assert.InRange(separator, begin + 1, text.Length - 1);
                Assert.InRange(finish, separator + 1, text.Length - 1);
                fields.Add((text.Substring(begin + 1, separator - begin - 1), text.Substring(separator + 1, finish - separator - 1)));
                begin = text.IndexOf('\u0013', finish + 1);
            }
            Assert.Equal(profile == "COUNTS" ? 30 : 10, fields.Count);
            string[] expectedInstructions = profile == "OMML" ? new[] { " EQ  \\* MERGEFORMAT " }
                : instructions.Select(instruction => " " + instruction + " ").ToArray();
            foreach (string instruction in expectedInstructions) {
                var matching = fields.Where(field => field.Instruction == instruction).ToArray();
                Assert.Equal(10, matching.Length);
                Assert.All(matching, field => Assert.Equal(profile == "COUNTS" ? "1" : string.Empty, field.Result));
            }

            byte[] emptyCounts = profile == "COUNTS" ? ClearCountCaches(encoded) : encoded;
            byte[][] importProfiles = profile == "COUNTS" ? new[] { encoded, emptyCounts, RemoveEmptyFieldSeparators(emptyCounts) } : new[] { encoded, RemoveEmptyFieldSeparators(encoded) };
            for (int importIndex = 0; importIndex < importProfiles.Length; importIndex++) {
                byte[] importProfile = importProfiles[importIndex];
                bool unchangedCounts = profile == "COUNTS" && importIndex == 0;
                int expectedLosses = profile == "COUNTS" ? 20 : 10;
                using WordDocument loaded = WordDocument.Load(new MemoryStream(importProfile));
                var findings = loaded.LegacyDocUnsupportedFeatures.Where(feature => feature.Code == "doc-field-empty-result-unprojected").ToArray();
                if (unchangedCounts) {
                    Assert.Empty(findings);
                    Assert.Empty(loaded.LegacyDocUnsupportedFeatures);
                    Assert.Equal(30, loaded.InspectFields().Count);
                } else {
                    Assert.Equal(expectedLosses, findings.Length);
                    Assert.Equal(expectedLosses, loaded.LegacyDocUnsupportedFeatures.Count);
                    Assert.Equal(expectedLosses, loaded.LegacyDocImportDiagnostics.Count(diagnostic => diagnostic.Code == "doc-field-empty-result-unprojected"));
                    Assert.DoesNotContain(loaded.LegacyDocUnsupportedFeatures, feature => feature.Code == "doc-field-lock-unprojected");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Main:EmptyResult");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Header:EmptyResult");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Footnote:EmptyResult");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Endnote:EmptyResult");
                    if (profile == "COUNTS") Assert.Equal(10, loaded.InspectFields().Count);
                    else Assert.Empty(loaded.InspectFields());
                    Assert.False(loaded.AssessLegacyDocWrite().IsSupported);
                }
                using var strict = new MemoryStream();
                if (unchangedCounts) loaded.Save(strict);
                else {
                    Assert.Throws<NotSupportedException>(() => loaded.Save(strict));
                    Assert.Equal(0, strict.Length);
                    loaded.Save(strict, new WordSaveOptions { AllowLoss = true });
                }
                Assert.True(strict.Length > 0);
            }
        } finally { DeleteIfExists(path); }
    }

    private static byte[] ClearCountCaches(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        byte[] table = compound.Streams["1Table"].ToArray();
        for (int index = BitConverter.ToInt32(main, 0x18); index < BitConverter.ToInt32(main, 0x1C); index++) {
            if (main[index] != 0x14) continue;
            Assert.Equal((byte)'1', main[index + 1]);
            Assert.Equal(0x15, main[index + 2]);
            main[index + 1] = 0x15;
            main[index + 2] = (byte)' ';
        }
        foreach (int directory in new[] { 0x11A, 0x122, 0x12A, 0x132, 0x21A }) {
            int length = BitConverter.ToInt32(main, directory + 4);
            if (length == 0) continue;
            int start = BitConverter.ToInt32(main, directory);
            int count = (length - 4) / 6;
            int properties = start + (count + 1) * 4;
            for (int index = 0; index < count; index++) {
                if ((table[properties + index * 2] & 0x1F) != 0x15) continue;
                int cp = BitConverter.ToInt32(table, start + index * 4);
                Buffer.BlockCopy(BitConverter.GetBytes(cp - 1), 0, table, start + index * 4, 4);
            }
        }
        return OfficeCompoundFileWriter.Write(compound.Streams.Select(stream => new OfficeCompoundStream(stream.Key,
            stream.Key == "WordDocument" ? main : stream.Key == "1Table" ? table : stream.Value)).ToArray());
    }

    private static byte[] RemoveEmptyFieldSeparators(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        byte[] table = compound.Streams["1Table"].ToArray();
        for (int index = BitConverter.ToInt32(main, 0x18); index < BitConverter.ToInt32(main, 0x1C); index++) {
            if (main[index] != 0x14) continue;
            Assert.Equal(0x15, main[index + 1]);
            main[index] = (byte)' ';
        }
        foreach (int directory in new[] { 0x11A, 0x122, 0x12A, 0x132, 0x21A }) {
            int length = BitConverter.ToInt32(main, directory + 4);
            if (length == 0) continue;
            int start = BitConverter.ToInt32(main, directory);
            int count = (length - 4) / 6;
            int properties = start + (count + 1) * 4;
            int terminal = BitConverter.ToInt32(table, start + count * 4);
            var records = new List<(int Cp, byte Character, byte Value)>();
            for (int index = 0; index < count; index++) {
                byte character = table[properties + index * 2];
                if ((character & 0x1F) == 0x14) continue;
                byte value = table[properties + index * 2 + 1];
                records.Add((BitConverter.ToInt32(table, start + index * 4), character,
                    (character & 0x1F) == 0x15 ? (byte)(value & ~0x80) : value));
            }
            for (int index = 0; index < records.Count; index++) Buffer.BlockCopy(BitConverter.GetBytes(records[index].Cp), 0, table, start + index * 4, 4);
            Buffer.BlockCopy(BitConverter.GetBytes(terminal), 0, table, start + records.Count * 4, 4);
            int outputProperties = start + (records.Count + 1) * 4;
            for (int index = 0; index < records.Count; index++) {
                table[outputProperties + index * 2] = records[index].Character;
                table[outputProperties + index * 2 + 1] = records[index].Value;
            }
            Buffer.BlockCopy(BitConverter.GetBytes(4 + records.Count * 6), 0, main, directory + 4, 4);
        }
        return OfficeCompoundFileWriter.Write(compound.Streams.Select(stream => new OfficeCompoundStream(stream.Key,
            stream.Key == "WordDocument" ? main : stream.Key == "1Table" ? table : stream.Value)).ToArray());
    }
}
