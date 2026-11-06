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
                paragraph.Append(new W.SimpleField { Instruction = " " + instruction + " ", FieldLock = true },
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

            byte[][] importProfiles = profile == "COUNTS" ? new[] { encoded } : new[] { encoded, RemoveEmptyFieldSeparators(encoded) };
            foreach (byte[] importProfile in importProfiles) {
                using WordDocument loaded = WordDocument.Load(new MemoryStream(importProfile));
                var findings = loaded.LegacyDocUnsupportedFeatures.Where(feature => feature.Code == "doc-field-empty-result-unprojected").ToArray();
                if (profile == "COUNTS") {
                    Assert.Empty(findings);
                    Assert.Empty(loaded.LegacyDocUnsupportedFeatures);
                    Assert.Equal(30, loaded.InspectFields().Count);
                } else {
                    Assert.Equal(10, findings.Length);
                    Assert.Equal(10, loaded.LegacyDocUnsupportedFeatures.Count);
                    Assert.Equal(10, loaded.LegacyDocImportDiagnostics.Count(diagnostic => diagnostic.Code == "doc-field-empty-result-unprojected"));
                    Assert.DoesNotContain(loaded.LegacyDocUnsupportedFeatures, feature => feature.Code == "doc-field-lock-unprojected");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Main:EmptyResult");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Header:EmptyResult");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Footnote:EmptyResult");
                    Assert.Contains(findings, feature => feature.DetailCode == "Field:Endnote:EmptyResult");
                    Assert.Empty(loaded.InspectFields());
                    Assert.False(loaded.AssessLegacyDocWrite().IsSupported);
                }
                using var strict = new MemoryStream();
                if (profile == "COUNTS") loaded.Save(strict);
                else {
                    Assert.Throws<NotSupportedException>(() => loaded.Save(strict));
                    Assert.Equal(0, strict.Length);
                    loaded.Save(strict, new WordSaveOptions { AllowLoss = true });
                }
                Assert.True(strict.Length > 0);
            }
        } finally { DeleteIfExists(path); }
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
