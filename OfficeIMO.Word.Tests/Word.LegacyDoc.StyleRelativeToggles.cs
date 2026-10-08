using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0x80, true, false)]
    [InlineData(0x81, false, true)]
    public void LegacyDoc_TextboxRelativeTogglesUseEachParagraphStyle(int operand, bool expectedStyled, bool expectedNormal) {
        const string body = "Body before textbox\r";
        string source = Path.Combine(_directoryWithFiles, $"TextboxRelative{operand}.doc");
        using (var document = WordDocument.Create()) {
            document.AddParagraph(body.TrimEnd('\r'));
            WordParagraph styled = document.AddParagraph("StyledTextboxMarker").SetStyle(WordParagraphStyles.Heading4);
            styled.Bold = true;
            styled.Italic = true;
            WordParagraph normalParagraph = document.AddParagraph("NormalTextboxMarker").SetStyle(WordParagraphStyles.Normal);
            normalParagraph.Bold = true;
            normalParagraph.Italic = true;
            Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style normal = styles.Elements<Style>().Single(style => style.StyleId == "Normal");
            normal.StyleRunProperties = new StyleRunProperties(new Bold { Val = false }, new Italic { Val = false });
            EnsureParagraphStyle(styles, WordParagraphStyles.Heading4.ToStringStyle()).StyleRunProperties = new StyleRunProperties(new Bold(), new Italic());
            document.Save(source);
        }
        byte[] bytes = RewriteNativeRunToggleOperands(File.ReadAllBytes(source), (byte)operand, minimumChanges: 2);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"];
        int originalBodyLength = BitConverter.ToInt32(word, 0x4C);
        // Reclassify the two native paragraphs as the body textbox story; retain their real PAPX/CHPX records.
        Buffer.BlockCopy(BitConverter.GetBytes(body.Length), 0, word, 0x4C, sizeof(int));
        Buffer.BlockCopy(BitConverter.GetBytes(originalBodyLength - body.Length), 0, word, 0x64, sizeof(int));
        using var package = new MemoryStream();
        using (RootStorage root = RootStorage.Create(package, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen)) {
            foreach (KeyValuePair<string, byte[]> entry in compound.Streams) {
                using CfbStream stream = root.CreateStream(entry.Key);
                stream.Write(entry.Value, 0, entry.Value.Length);
            }
        }
        using WordDocument loaded = WordDocument.Load(new MemoryStream(package.ToArray()));
        AssertTextbox(loaded);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(loaded.ToBytes()));
        AssertTextbox(reopened);

        void AssertTextbox(WordDocument document) {
            WordTextBox box = Assert.Single(document.TextBoxes);
            Paragraph[] paragraphs = box.Content!.Elements<Paragraph>().ToArray();
            Assert.Equal(2, paragraphs.Length);
            var styled = new WordParagraph(document, paragraphs[0], newRun: false);
            var normal = new WordParagraph(document, paragraphs[1], newRun: false);
            WordParagraph styledRun = Assert.Single(styled.GetRuns());
            WordParagraph normalRun = Assert.Single(normal.GetRuns());
            Assert.Equal("StyledTextboxMarker", styledRun.Text);
            Assert.Equal("NormalTextboxMarker", normalRun.Text);
            Assert.Equal(expectedStyled, styledRun.Bold);
            Assert.Equal(expectedStyled, styledRun.Italic);
            Assert.Equal(expectedNormal, normalRun.Bold);
            Assert.Equal(expectedNormal, normalRun.Italic);
            Assert.Equal(WordParagraphStyles.Heading4, styled.Style);
        }
    }

    [Theory]
    [InlineData(0, false, false)]
    [InlineData(0, true, false)]
    [InlineData(1, false, true)]
    [InlineData(1, true, true)]
    [InlineData(0x80, false, false)]
    [InlineData(0x80, true, true)]
    [InlineData(0x81, false, true)]
    [InlineData(0x81, true, false)]
    public void LegacyDoc_StyleRelativeTogglesResolveBeforeBodyTableHeaderAndNoteRuns(int operand, bool inherited, bool expected) {
        string source = Path.Combine(_directoryWithFiles, $"RelativeToggles{operand}_{inherited}.doc");
        using (var document = WordDocument.Create()) {
            WordParagraph body = document.AddParagraph("BodyToggleMarker");
            Seed(body);
            Seed(document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("TableToggleMarker", removeExistingParagraphs: true));
            document.AddHeadersAndFooters();
            Seed(document.Sections[0].Header.Default!.AddParagraph("HeaderToggleMarker"));
            Seed(document.Sections[0].Footer.Default!.AddParagraph("FooterToggleMarker"));
            WordParagraph footnote = body.AddFootNote("FootnoteToggleMarker").FootNote!.Paragraphs![1];
            Seed(footnote);
            Seed(footnote.AddParagraph("LaterFootnoteToggleMarker"));
            WordParagraph endnote = body.AddEndNote("EndnoteToggleMarker").EndNote!.Paragraphs![1];
            Seed(endnote);
            Seed(endnote.AddParagraph("LaterEndnoteToggleMarker"));
            body.AddComment("OfficeIMO", "OI", "CommentToggleMarker");
            Seed(document.Comments.Single().Paragraphs.Single());
            Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style normal = styles.Elements<Style>().Single(style => style.StyleId == "Normal");
            normal.StyleRunProperties = new StyleRunProperties(new Bold { Val = inherited }, new Italic { Val = inherited });
            Style heading = EnsureParagraphStyle(styles, WordParagraphStyles.Heading4.ToStringStyle());
            heading.BasedOn = new BasedOn { Val = "Normal" };
            heading.StyleRunProperties = new StyleRunProperties();
            document.Save(source);

            static void Seed(WordParagraph paragraph) {
                paragraph.SetStyle(WordParagraphStyles.Heading4);
                paragraph.Bold = true;
                paragraph.Italic = true;
            }
        }

        byte[] bytes = RewriteNativeRunToggleOperands(File.ReadAllBytes(source), (byte)operand);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
        WordSection section = loaded.Sections[0];
        WordParagraph[] markers = loaded.Paragraphs
            .Concat(loaded.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs))
            .Concat(section.Header.Default!.Paragraphs)
            .Concat(section.Footer.Default!.Paragraphs)
            .Concat(loaded.FootNotes.SelectMany(note => note.Paragraphs!))
            .Concat(loaded.EndNotes.SelectMany(note => note.Paragraphs!))
            .Concat(loaded.Comments.SelectMany(comment => comment.Paragraphs))
            .Where(paragraph => paragraph.Text.EndsWith("ToggleMarker", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(9, markers.Length);
        foreach (WordParagraph marker in markers) {
            Assert.Equal(expected, marker.Bold);
            Assert.Equal(expected, marker.Italic);
        }
        using WordDocument reopened = WordDocument.Load(new MemoryStream(loaded.ToBytes()));
        WordParagraph bodyMarker = Assert.Single(reopened.Paragraphs, paragraph => paragraph.Text == "BodyToggleMarker");
        Assert.Equal(expected, bodyMarker.Bold);
        Assert.Equal(expected, bodyMarker.Italic);
    }

    [Fact]
    public void LegacyDoc_StyleRelativeToggleOperandsPreserveAllFlagsAndLastDeclaration() {
        ushort[] sprms = { 0x0835, 0x0836, 0x0837, 0x2A53, 0x0838, 0x0839, 0x0858, 0x0854, 0x083C, 0x0875, 0x083B, 0x083A };
        LegacyDocCharacterFormatProperties[] properties = {
            LegacyDocCharacterFormatProperties.Bold, LegacyDocCharacterFormatProperties.Italic,
            LegacyDocCharacterFormatProperties.Strike, LegacyDocCharacterFormatProperties.DoubleStrike,
            LegacyDocCharacterFormatProperties.Outline, LegacyDocCharacterFormatProperties.Shadow,
            LegacyDocCharacterFormatProperties.Emboss, LegacyDocCharacterFormatProperties.Imprint,
            LegacyDocCharacterFormatProperties.Hidden, LegacyDocCharacterFormatProperties.NoProof,
            LegacyDocCharacterFormatProperties.Caps, LegacyDocCharacterFormatProperties.SmallCaps
        };
        for (int index = 0; index < sprms.Length; index++) {
            LegacyDocCharacterFormat match = Read(sprms[index], 0x80);
            LegacyDocCharacterFormat invert = Read(sprms[index], 0x81);
            Assert.Equal(properties[index], match.StyleRelative);
            Assert.Equal(LegacyDocCharacterFormatProperties.None, match.StyleInverted);
            Assert.Equal(properties[index], invert.StyleRelative);
            Assert.Equal(properties[index], invert.StyleInverted);
            Assert.NotEqual(match, invert);
            Assert.NotEqual(match, Read(sprms[index], 0));
            Assert.NotEqual(invert, Read(sprms[index], 1));
            Assert.Equal(match.StyleRelative, match.WithDefaultFont("Arial").StyleRelative);
            Assert.Equal(invert.StyleInverted, invert.WithDefaultFont("Arial").StyleInverted);
            LegacyDocTextRun field = LegacyDocTextRunFactory.CreateFieldRun("1", LegacyDocFieldKind.Page, invert, new[] { 0 });
            Assert.Equal(invert.StyleRelative, field.StyleRelative);
            Assert.Equal(invert.StyleInverted, field.StyleInverted);
        }
        byte[] sequence = { 0x35, 0x08, 0x81, 0x35, 0x08, 0x80, 0x36, 0x08, 0x80, 0x36, 0x08, 0 };
        LegacyDocCharacterFormat last = LegacyDocCharacterFormattingReader.ReadGrpprl(sequence, 0, sequence.Length, Array.Empty<string>());
        Assert.Equal(LegacyDocCharacterFormatProperties.Bold, last.StyleRelative);
        Assert.Equal(LegacyDocCharacterFormatProperties.None, last.StyleInverted);
        Assert.True(last.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
        Assert.False(last.Italic);

        static LegacyDocCharacterFormat Read(ushort sprm, byte operand) {
            byte[] bytes = { (byte)sprm, (byte)(sprm >> 8), operand };
            return LegacyDocCharacterFormattingReader.ReadGrpprl(bytes, 0, bytes.Length, Array.Empty<string>());
        }
    }

    private static byte[] RewriteNativeRunToggleOperands(byte[] source, byte operand, int minimumChanges = 16) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"];
        int plc = BitConverter.ToInt32(word, 0xFA);
        int bins = (BitConverter.ToInt32(word, 0xFE) - 4) / 8;
        int bte = plc + ((bins + 1) * 4);
        var changed = new HashSet<int>();
        for (int bin = 0; bin < bins; bin++) {
            int page = BitConverter.ToInt32(table, bte + (bin * 4)) * 512;
            int runs = word[page + 511];
            int offsets = page + ((runs + 1) * 4);
            for (int run = 0; run < runs; run++) {
                int relative = word[offsets + run] * 2;
                if (relative == 0) continue;
                int start = page + relative + 1;
                int end = start + word[page + relative];
                for (int position = start; position + 2 < end; position++) {
                    if (word[position + 1] != 0x08 || (word[position] != 0x35 && word[position] != 0x36)) continue;
                    if (!changed.Add(position + 2)) continue;
                    Assert.Equal(1, word[position + 2]);
                    word[position + 2] = operand;
                }
            }
        }
        Assert.True(changed.Count >= minimumChanges);
        using var output = new MemoryStream();
        using (RootStorage root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen)) {
            foreach (KeyValuePair<string, byte[]> entry in compound.Streams) {
                using CfbStream stream = root.CreateStream(entry.Key);
                stream.Write(entry.Value, 0, entry.Value.Length);
            }
        }
        return output.ToArray();
    }
}
