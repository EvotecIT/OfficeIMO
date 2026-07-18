namespace OfficeIMO.OneNote.Tests;

public sealed class CorruptionSafetyTests {
    public static IEnumerable<object[]> Fixtures() {
        yield return new object[] { "testOneNote.one" };
        yield return new object[] { "testOneNote2016.one" };
        yield return new object[] { "testOneNoteEmbeddedWordDoc.one" };
        yield return new object[] { "testOneNoteFromOffice365.one" };
        yield return new object[] { "testOneNoteFromOffice365-2.one" };
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void DeterministicByteMutationsEitherParseOrFailWithBoundedIoErrors(string fixture) {
        byte[] source = File.ReadAllBytes(FixturePath(fixture));
        OneNoteReaderOptions options = TightOptions(source.Length);

        for (int sample = 0; sample < 24; sample++) {
            byte[] mutated = (byte[])source.Clone();
            int offset = sample * (source.Length - 1) / 23;
            mutated[offset] ^= (byte)(0x5A ^ sample);

            try {
                OneNoteSection section = OneNoteSectionReader.Read(new MemoryStream(mutated), options);
                Assert.NotNull(section);
            } catch (IOException) {
                // Stable bounded parse failures are expected for malformed mutations.
            } catch (Exception exception) {
                Assert.Fail("Mutation at offset " + offset + " escaped the bounded parser as " + exception.GetType().FullName + ": " + exception.Message);
            }
        }
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void TruncationSweepEitherParsesOrFailsWithBoundedIoErrors(string fixture) {
        byte[] source = File.ReadAllBytes(FixturePath(fixture));
        int[] lengths = { 0, 1, 47, Math.Min(512, source.Length - 1), source.Length / 2, source.Length - 1 };

        foreach (int length in lengths.Distinct().Where(value => value >= 0 && value < source.Length)) {
            var truncated = new byte[length];
            Buffer.BlockCopy(source, 0, truncated, 0, length);
            Exception? exception = Record.Exception(() => OneNoteSectionReader.Read(new MemoryStream(truncated), TightOptions(source.Length)));

            if (exception != null) Assert.IsAssignableFrom<IOException>(exception);
        }
    }

    [Fact]
    public void WriterHonorsOutputBoundBeforeReturningAnArtifact() {
        var section = new OneNoteSection { Name = "Bounded" };
        var page = new OneNotePage { Title = "Page" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Bounded content" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder(1024).BuildSection(section);
        IOException layoutException = Assert.Throws<IOException>(() =>
            OneNoteRevisionStoreWriter.Write(graph, 1024));

        IOException exception = Assert.Throws<IOException>(() =>
            OneNoteSectionWriter.Write(section, new OneNoteWriterOptions { MaxOutputBytes = 1024 }));
        IOException packageException = Assert.Throws<IOException>(() =>
            OneNoteSectionWriter.Write(section, new OneNoteWriterOptions {
                StorageFormat = OneNoteStorageFormat.FileSynchronizationPackage,
                MaxOutputBytes = 1024
            }));

        Assert.Contains("MaxOutputBytes", layoutException.Message);
        Assert.Contains("MaxOutputBytes", exception.Message);
        Assert.Contains("MaxOutputBytes", packageException.Message);
    }

    [Fact]
    public void WriterRoundTripValidationAllowsAssetsWithinTheOutputBound() {
        long outputLimit = OneNoteReaderOptions.DefaultMaxAssetBytes + 1;
        var writerOptions = new OneNoteWriterOptions {
            MaxPageRelationshipDepth = 192,
            MaxContentDepth = 160,
            MaxInkPathValues = 1_250_000
        };

        OneNoteReaderOptions options = OneNoteWriterValidation.CreateReaderOptions(writerOptions, outputLimit);

        Assert.Equal(outputLimit, options.MaxInputBytes);
        Assert.Equal(outputLimit, options.MaxAssetBytes);
        Assert.Equal(outputLimit, options.MaxTotalAssetBytes);
        Assert.Equal(writerOptions.MaxPageRelationshipDepth, options.MaxPageRelationshipDepth);
        Assert.Equal(writerOptions.MaxContentDepth, options.MaxPropertySetDepth);
        Assert.Equal(writerOptions.MaxInkPathValues, options.MaxInkPathValues);
    }

    [Fact]
    public void OversizedTextRunBoundariesAreBoundedFormatErrors() {
        var section = new OneNoteSection { Name = "Malformed runs" };
        var page = new OneNotePage { Title = "Page" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "first" });
        paragraph.Runs.Add(new OneNoteTextRun { Text = "second" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces[1];
        int richTextIndex = pageSpace.Objects.ToList().FindIndex(item => item.Jcid == OneNoteSchema.JcidRichTextNode);
        OneNoteWriteObject richText = pageSpace.Objects[richTextIndex];
        OneNoteWriteProperty[] properties = richText.Properties
            .Select(property => (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.TextRunIndex
                ? new OneNoteWriteProperty(property.RawId, data: BitConverter.GetBytes(uint.MaxValue), preserveRawId: true)
                : property)
            .ToArray();
        pageSpace.Objects[richTextIndex] = new OneNoteWriteObject(richText.Id, richText.Jcid, properties);
        byte[] data = OneNoteRevisionStoreWriter.Write(graph);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionReader.Read(new MemoryStream(data)));

        Assert.Equal("ONENOTE_TEXT_RUN_BOUNDARY", exception.Code);
    }

    private static OneNoteReaderOptions TightOptions(int inputLength) => new OneNoteReaderOptions {
        MaxInputBytes = Math.Max(1, inputLength),
        MaxFileNodeListFragments = 1024,
        MaxFileNodes = 10000,
        MaxTransactionLogFragments = 1024,
        MaxTransactionEntries = 10000,
        MaxObjects = 10000,
        MaxPropertiesPerObject = 4096,
        MaxPropertySetDepth = 32,
        MaxAssetBytes = 1024 * 1024,
        MaxTotalAssetBytes = 1024 * 1024,
        MaxStreamObjects = 10000,
        MaxStreamObjectDepth = 32
    };

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);
}
