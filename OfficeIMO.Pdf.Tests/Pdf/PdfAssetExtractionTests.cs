using System.Buffers.Binary;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfAssetExtractionTests {
    [Fact]
    public void BundleExportsDeduplicatedFilesAndPlacementAndPrintEvidence() {
        byte[] icc = PdfIccProfiles.SrgbIec6196621;
        using PdfDocument document = PdfDocument.Load(BuildPdf(icc));
        PdfImagePlacement softMaskPlacement = Assert.Single(document.Reader.ImagePlacements(), p => p.ResourceName == "MaskImage");
        Assert.Equal(3D, softMaskPlacement.EffectiveDpiX);
        Assert.Equal(3D, softMaskPlacement.EffectiveDpiY);

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-assets-" + Guid.NewGuid().ToString("N"));
        try {
            IReadOnlyList<string> files = document.ExtractAssets(directory);
            Assert.Equal(2, files.Count(name => name.StartsWith("images/", StringComparison.Ordinal)));
            Assert.EndsWith(".ttf", Assert.Single(files, name => name.StartsWith("fonts/", StringComparison.Ordinal)));
            Assert.EndsWith(".icc", Assert.Single(files, name => name.StartsWith("profiles/", StringComparison.Ordinal)));
            Assert.EndsWith(".svg", Assert.Single(files, name => name.StartsWith("pages/", StringComparison.Ordinal)));

            using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(directory, "report.json")));
            JsonElement root = json.RootElement;
            Assert.Equal("officeimo.pdf.assets.v1", root.GetProperty("schema").GetString());
            JsonElement[] images = root.GetProperty("images").EnumerateArray().ToArray();
            JsonElement softMask = Assert.Single(images, image => image.GetProperty("resource").GetString() == "MaskImage");
            Assert.True(softMask.GetProperty("maskResolved").GetBoolean());
            Assert.True(softMask.GetProperty("rgbaPng").GetBoolean());
            Assert.Equal(2, softMask.GetProperty("widthPixels").GetInt32());
            Assert.Equal(1, softMask.GetProperty("heightPixels").GetInt32());
            byte[] png = ReadFile(softMask);
            Assert.Equal(8, png[24]);
            Assert.Equal(6, png[25]);
            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
            Assert.Equal(new byte[] { 32, 32, 32, 0, 200, 200, 200, 128 }, Assert.IsType<OfficeRasterImage>(raster).GetPixels());
            JsonElement[] iccImages = images.Where(image => image.GetProperty("colorSpace").GetString() == "ICCBased").ToArray();
            Assert.Equal(2, iccImages.Length);
            Assert.Equal(FileOf(iccImages[0]), FileOf(iccImages[1]));
            Assert.NotEqual(iccImages[0].GetProperty("objectNumber").GetInt32(), iccImages[1].GetProperty("objectNumber").GetInt32());

            JsonElement placement = Assert.Single(root.GetProperty("imagePlacements").EnumerateArray(),
                item => item.GetProperty("resource").GetString() == "MaskImage");
            Assert.Equal(FileOf(softMask), FileOf(placement));
            Assert.Equal(3D, placement.GetProperty("effectiveDpiX").GetDouble());
            Assert.Equal(3D, placement.GetProperty("effectiveDpiY").GetDouble());

            JsonElement trim = Assert.Single(root.GetProperty("pageBoxes").EnumerateArray(), box => box.GetProperty("box").GetString() == "TrimBox");
            Assert.Equal(50.8D, trim.GetProperty("widthMm").GetDouble(), 4);
            Assert.Equal(25.4D, trim.GetProperty("heightMm").GetDouble(), 4);
            JsonElement ink = Assert.Single(root.GetProperty("spotColorants").EnumerateArray());
            Assert.Equal("Separation", ink.GetProperty("colorSpace").GetString());
            Assert.Equal("Brand<Blue>", ink.GetProperty("name").GetString());
            JsonElement[] profiles = root.GetProperty("iccProfiles").EnumerateArray().ToArray();
            Assert.Equal(2, profiles.Length);
            Assert.Contains(profiles, profile => profile.GetProperty("source").GetString() == "OutputIntent");
            Assert.Contains(profiles, profile => profile.GetProperty("source").GetString() == "ICCBased");
            Assert.Equal(FileOf(profiles[0]), FileOf(profiles[1]));
            Assert.Equal(icc, ReadFile(profiles[0]));

            JsonElement[] fonts = root.GetProperty("fonts").EnumerateArray().ToArray();
            JsonElement[] embedded = fonts.Where(font => font.GetProperty("embedded").GetBoolean()).ToArray();
            Assert.Equal(2, embedded.Length);
            Assert.All(embedded, font => {
                Assert.True(font.GetProperty("subset").GetBoolean());
                Assert.Equal("ABCDEF", font.GetProperty("subsetTag").GetString());
                Assert.True(font.GetProperty("nameTableRepaired").GetBoolean());
            });
            Assert.Contains(fonts, font => !font.GetProperty("embedded").GetBoolean() && font.GetProperty("file").ValueKind == JsonValueKind.Null);
            Assert.Equal(FileOf(embedded[0]), FileOf(embedded[1]));
            byte[] ttf = ReadFile(embedded[0]);
            OfficeTrueTypeFont face = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(ttf));
            Assert.True(face.MatchesFamilyName("AssetFixture ABCDEF"));
            Assert.NotEmpty(face.GetTextContours("A", 0, 0, 10));
            Assert.Equal(0xB1B0AFBAU, SfntChecksum(ttf));

            JsonElement page = Assert.Single(root.GetProperty("pages").EnumerateArray());
            Assert.Equal(XName.Get("svg", "http://www.w3.org/2000/svg"), XDocument.Load(Path.Combine(directory, FileOf(page))).Root!.Name);
            string html = File.ReadAllText(Path.Combine(directory, "report.html"));
            Assert.Contains("href=\"" + FileOf(softMask) + "\"", html);
            Assert.Contains("Brand&lt;Blue&gt;", html);

            string fontDirectory = Path.Combine(directory, "standalone");
            string standaloneFont = Assert.Single(document.ExtractEmbeddedFonts(fontDirectory));
            Assert.Equal(ttf, File.ReadAllBytes(Path.Combine(fontDirectory, standaloneFont)));
            Assert.Throws<IOException>(() => document.ExtractAssets(directory));

            byte[] ReadFile(JsonElement item) => File.ReadAllBytes(Path.Combine(directory, FileOf(item)));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void StreamFactoryClosesEachEntryAndCanUseOrdinarySoftMaskExport() {
        using PdfDocument document = PdfDocument.Load(BuildPdf(PdfIccProfiles.SrgbIec6196621));
        var streams = new Dictionary<string, MemoryStream>(StringComparer.Ordinal);
        var options = new PdfAssetExtractionOptions { CompositeSoftMasks = false };
        IReadOnlyList<string> paths = document.ExtractAssetsTo(path => {
            Assert.All(streams.Values, stream => Assert.False(stream.CanWrite));
            var stream = new MemoryStream();
            streams.Add(path, stream);
            return stream;
        }, options);

        Assert.Equal(paths.Count, streams.Count);
        Assert.All(streams.Values, stream => Assert.False(stream.CanWrite));
        using JsonDocument json = JsonDocument.Parse(streams["report.json"].ToArray());
        JsonElement image = Assert.Single(json.RootElement.GetProperty("images").EnumerateArray(),
            item => item.GetProperty("resource").GetString() == "MaskImage");
        Assert.True(image.GetProperty("maskResolved").GetBoolean());
        Assert.False(image.GetProperty("rgbaPng").GetBoolean());
        Assert.Equal(4, streams[FileOf(image)].ToArray()[25]);
    }

    [Fact]
    public void BundleResolvesSharedImagesInTheirFormContextsAndInventoriesAnnotationImages() {
        using PdfDocument document = PdfDocument.Load(BuildVariantPdf());
        var streams = new Dictionary<string, MemoryStream>(StringComparer.Ordinal);
        document.ExtractAssetsTo(path => streams[path] = new MemoryStream());
        using JsonDocument json = JsonDocument.Parse(streams["report.json"].ToArray());
        JsonElement root = json.RootElement;
        JsonElement[] placements = root.GetProperty("imagePlacements").EnumerateArray().ToArray();
        Assert.Equal(2, placements.Length);
        Assert.Equal(placements[0].GetProperty("objectNumber").GetInt32(), placements[1].GetProperty("objectNumber").GetInt32());
        Assert.NotEqual(FileOf(placements[0]), FileOf(placements[1]));
        Assert.True(OfficePngReader.TryDecode(streams[FileOf(placements[0])].ToArray(), out OfficeRasterImage? gray));
        Assert.True(OfficePngReader.TryDecode(streams[FileOf(placements[1])].ToArray(), out OfficeRasterImage? red));
        Assert.Equal(new byte[] { 255, 255, 255, 255 }, Assert.IsType<OfficeRasterImage>(gray).GetPixels().Take(4));
        Assert.Equal(new byte[] { 255, 0, 0, 255 }, Assert.IsType<OfficeRasterImage>(red).GetPixels().Take(4));

        var svg = XDocument.Parse(Encoding.UTF8.GetString(streams["pages/0001.svg"].ToArray()));
        byte[][] svgColors = svg.Descendants(XName.Get("image", "http://www.w3.org/2000/svg"))
            .Select(image => (string?)image.Attribute("href") ?? (string?)image.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")))
            .OfType<string>()
            .Where(uri => uri.StartsWith("data:image/png;base64,", StringComparison.Ordinal))
            .Select(uri => Convert.FromBase64String(uri.Substring("data:image/png;base64,".Length)))
            .Select(bytes => {
                Assert.True(OfficePngReader.TryDecode(bytes, out OfficeRasterImage? pixels));
                return Assert.IsType<OfficeRasterImage>(pixels).GetPixels();
            }).ToArray();
        Assert.Contains(svgColors, colors => colors.Take(4).SequenceEqual(new byte[] { 255, 255, 255, 255 }));
        Assert.Contains(svgColors, colors => colors.Take(4).SequenceEqual(new byte[] { 255, 0, 0, 255 }));

        JsonElement stamp = Assert.Single(root.GetProperty("images").EnumerateArray(),
            image => image.GetProperty("resource").GetString() == "StampImage");
        Assert.True(stamp.GetProperty("imageFile").GetBoolean());
        Assert.True(streams.ContainsKey(FileOf(stamp)));
        Assert.Contains(root.GetProperty("diagnostics").EnumerateArray(),
            item => item.GetProperty("code").GetString() == "AdditionalImagePlacementContexts");
    }

    [Fact]
    public void FailedFolderExportDoesNotPublishPartialFiles() {
        using PdfDocument document = PdfDocument.Load(BuildPdf(PdfIccProfiles.SrgbIec6196621));
        string parent = Path.Combine(Path.GetTempPath(), "officeimo-asset-limit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(parent);
        try {
            Assert.Throws<InvalidDataException>(() =>
                document.ExtractAssets(Path.Combine(parent, "output"), new PdfAssetExtractionOptions { MaximumOutputBytes = 1 }));
            Assert.Empty(Directory.EnumerateFileSystemEntries(parent));
        } finally {
            Directory.Delete(parent, recursive: true);
        }
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("bad-offset")]
    [InlineData("null-names")]
    public void FontFile3ExportsOpenTypeWithARebuiltNameTable(string damage) {
        string fontPath = Path.Combine(AppContext.BaseDirectory, "OfficeIMO.Pdf.Tests", "Pdf", "Fixtures", "Fonts", "SourceSerif4-Regular.otf");
        byte[] font = File.ReadAllBytes(fontPath);
        int tables = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4, 2));
        int nameRecord = Assert.Single(Enumerable.Range(0, tables).Select(i => 12 + 16 * i),
            offset => Encoding.ASCII.GetString(font, offset, 4) == "name");
        switch (damage) {
            case "missing":
                Encoding.ASCII.GetBytes("zzzz").CopyTo(font, nameRecord);
                break;
            case "bad-offset":
                BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(nameRecord + 8, 4), uint.MaxValue);
                break;
            case "null-names":
                int body = checked((int)BinaryPrimitives.ReadUInt32BigEndian(font.AsSpan(nameRecord + 8, 4)));
                int length = checked((int)BinaryPrimitives.ReadUInt32BigEndian(font.AsSpan(nameRecord + 12, 4)));
                int strings = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(body + 4, 2));
                Array.Clear(font, body + strings, length - strings);
                break;
        }
        using PdfDocument document = PdfDocument.Load(BuildPdf(PdfIccProfiles.SrgbIec6196621, font));
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-otf-assets-" + Guid.NewGuid().ToString("N"));
        try {
            string file = Assert.Single(document.ExtractEmbeddedFonts(directory));
            Assert.EndsWith(".otf", file);
            byte[] repaired = File.ReadAllBytes(Path.Combine(directory, file));
            Assert.Equal("OTTO", Encoding.ASCII.GetString(repaired, 0, 4));
            Assert.Equal(0xB1B0AFBAU, SfntChecksum(repaired));
            int count = BinaryPrimitives.ReadUInt16BigEndian(repaired.AsSpan(4, 2));
            int name = Assert.Single(Enumerable.Range(0, count).Select(i => 12 + 16 * i),
                offset => Encoding.ASCII.GetString(repaired, offset, 4) == "name");
            int length = checked((int)BinaryPrimitives.ReadUInt32BigEndian(repaired.AsSpan(name + 12, 4)));
            int start = checked((int)BinaryPrimitives.ReadUInt32BigEndian(repaired.AsSpan(name + 8, 4)));
            int strings = BinaryPrimitives.ReadUInt16BigEndian(repaired.AsSpan(start + 4, 2));
            Assert.Contains("AssetFixture ABCDEF", Encoding.BigEndianUnicode.GetString(repaired, start + strings, length - strings));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    private static string FileOf(JsonElement item) => item.GetProperty("file").GetString()!;

    private static uint SfntChecksum(byte[] bytes) {
        uint sum = 0;
        for (int offset = 0; offset < bytes.Length; offset += 4) {
            sum = unchecked(sum + BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4)));
        }
        return sum;
    }

    private static byte[] BuildPdf(byte[] icc, byte[]? openType = null) {
        byte[] font = openType ?? ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A');
        byte[] seed = PdfDocument.Create().Paragraph(paragraph => paragraph.AddText("A")).ToBytes();
        return PdfDocumentObjectGraphRewriter.Rewrite(seed, null, null, (objects, security) => {
            int next = objects.Keys.Max() + 1;
            PdfDictionary page = Assert.Single(objects.Values.Select(value => value.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            PdfDictionary catalog = Assert.Single(objects.Values.Select(value => value.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Catalog");
            PdfDictionary programDictionary = openType is null
                ? Dict(("Length1", new PdfNumber(font.Length)))
                : Dict(("Subtype", Name("OpenType")));
            PdfReference program = Add(new PdfStream(programDictionary, font));
            PdfReference descriptor = Add(Dict(("Type", Name("FontDescriptor")), ("FontName", Name("ABCDEF+AssetFixture")),
                ("Flags", new PdfNumber(32)), ("FontBBox", Numbers(0, 0, 600, 700)), ("ItalicAngle", new PdfNumber(0)),
                ("Ascent", new PdfNumber(700)), ("Descent", new PdfNumber(0)), ("StemV", new PdfNumber(80)),
                ("CapHeight", new PdfNumber(700)), (openType is null ? "FontFile2" : "FontFile3", program)));
            PdfReference primaryFont = Add(Font("ABCDEF+AssetFixture"));
            PdfReference duplicateFont = Add(Font("ABCDEF+FixtureAlias"));
            PdfReference plainFont = Add(Dict(("Type", Name("Font")), ("Subtype", Name("Type1")), ("BaseFont", Name("Helvetica"))));

            PdfReference profile = Add(new PdfStream(Dict(("N", new PdfNumber(3))), icc));
            catalog.Items["OutputIntents"] = ArrayOf(Add(Dict(("Type", Name("OutputIntent")), ("S", Name("GTS_PDFX")),
                ("OutputConditionIdentifier", new PdfStringObj("Fixture sRGB")), ("DestOutputProfile", profile))));
            PdfReference mask = AddWithIndirectDimensions(Image(Name("DeviceGray"), new byte[] { 0, 128 }));
            PdfReference masked = AddWithIndirectDimensions(Image(Name("DeviceGray"), new byte[] { 32, 200 }, mask));
            PdfArray iccSpace = ArrayOf(Name("ICCBased"), profile);
            byte[] rgb = { 20, 60, 220, 230, 110, 10 };
            PdfReference iccImage = Add(Image(iccSpace, rgb));
            PdfReference copy = Add(Image(iccSpace, rgb));
            PdfArray spot = ArrayOf(Name("Separation"), Name("Brand<Blue>"), Name("DeviceRGB"),
                Dict(("FunctionType", new PdfNumber(2)), ("Domain", Numbers(0, 1)), ("C0", Numbers(1, 1, 1)),
                    ("C1", Numbers(0, 0.2D, 0.8D)), ("N", new PdfNumber(1))));
            page.Items["MediaBox"] = Numbers(0, 0, 144, 144);
            page.Items["TrimBox"] = Numbers(5, 5, 77, 41);
            page.Items["UserUnit"] = new PdfNumber(2);
            page.Items["Resources"] = Dict(
                ("Font", Dict(("F1", primaryFont), ("F2", duplicateFont), ("Plain", plainFont))),
                ("XObject", Dict(("MaskImage", masked), ("ProfileImage", iccImage), ("ProfileCopy", copy))),
                ("ColorSpace", Dict(("Ink", spot))));
            page.Items["Contents"] = Add(new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes(
                "BT /F1 12 Tf 5 100 Td (A) Tj ET\n" +
                "q 0 24 -12 0 20 20 cm /MaskImage Do Q\n" +
                "q 24 0 0 12 40 20 cm /ProfileImage Do Q\n" +
                "q 24 0 0 12 70 20 cm /ProfileCopy Do Q\n" +
                "/Ink cs 0.5 sc 5 55 10 10 re f")));
            return security.InfoObjectNumber;

            PdfReference Add(PdfObject value) {
                int id = next++;
                objects.Add(id, new PdfIndirectObject(id, 0, value));
                return new PdfReference(id, 0);
            }

            PdfReference AddWithIndirectDimensions(PdfStream image) {
                foreach (string key in new[] { "Width", "Height", "BitsPerComponent" }) {
                    image.Dictionary.Items[key] = Add(image.Dictionary.Items[key]);
                }
                return Add(image);
            }

            PdfDictionary Font(string name) => Dict(("Type", Name("Font")), ("Subtype", Name(openType is null ? "TrueType" : "Type1")),
                ("BaseFont", Name(name)), ("FirstChar", new PdfNumber(65)), ("LastChar", new PdfNumber(65)),
                ("Widths", Numbers(600)), ("Encoding", Name("WinAnsiEncoding")), ("FontDescriptor", descriptor));
        });
    }

    private static byte[] BuildVariantPdf() {
        byte[] seed = PdfDocument.Create().Paragraph(paragraph => paragraph.AddText("A")).ToBytes();
        return PdfDocumentObjectGraphRewriter.Rewrite(seed, null, null, (objects, security) => {
            int next = objects.Keys.Max() + 1;
            PdfDictionary page = Assert.Single(objects.Values.Select(value => value.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            PdfReference shared = Add(Image(Name("Ink"), new byte[] { 255, 64 }));
            PdfReference stamp = Add(Image(Name("DeviceRGB"), new byte[] { 0, 90, 30, 0, 100, 5 }));
            PdfArray red = ArrayOf(Name("Separation"), Name("RedInk"), Name("DeviceRGB"),
                Dict(("FunctionType", new PdfNumber(2)), ("Domain", Numbers(0, 1)), ("C0", Numbers(1, 1, 1)),
                    ("C1", Numbers(1, 0, 0)), ("N", new PdfNumber(1))));
            PdfReference grayForm = Form("Shared", shared, Dict(("Ink", Name("DeviceGray"))));
            PdfReference redForm = Form("Shared", shared, Dict(("Ink", red)));
            PdfReference appearance = Form("StampImage", stamp, new PdfDictionary());
            page.Items["Annots"] = ArrayOf(Add(Dict(("Type", Name("Annot")), ("Subtype", Name("Stamp")),
                ("Rect", Numbers(60, 20, 80, 40)), ("AP", Dict(("N", appearance))))));
            page.Items["Resources"] = Dict(("XObject", Dict(("Gray", grayForm), ("Red", redForm))));
            page.Items["Contents"] = Add(new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes(
                "q 1 0 0 1 10 20 cm /Gray Do Q q 1 0 0 1 35 20 cm /Red Do Q")));
            return security.InfoObjectNumber;

            PdfReference Form(string name, PdfReference image, PdfDictionary colorSpaces) => Add(new PdfStream(
                Dict(("Type", Name("XObject")), ("Subtype", Name("Form")), ("BBox", Numbers(0, 0, 20, 20)),
                    ("Resources", Dict(("XObject", Dict((name, image))), ("ColorSpace", colorSpaces)))),
                Encoding.ASCII.GetBytes("q 20 0 0 20 0 0 cm /" + name + " Do Q")));

            PdfReference Add(PdfObject value) {
                int id = next++;
                objects.Add(id, new PdfIndirectObject(id, 0, value));
                return new PdfReference(id, 0);
            }
        });
    }

    private static PdfStream Image(PdfObject colorSpace, byte[] pixels, PdfReference? mask = null) {
        PdfDictionary dictionary = Dict(("Type", Name("XObject")), ("Subtype", Name("Image")), ("Width", new PdfNumber(2)),
            ("Height", new PdfNumber(1)), ("BitsPerComponent", new PdfNumber(8)), ("ColorSpace", colorSpace));
        if (mask is not null) dictionary.Items["SMask"] = mask;
        return new PdfStream(dictionary, pixels);
    }

    private static PdfDictionary Dict(params (string Key, PdfObject Value)[] entries) {
        var dictionary = new PdfDictionary();
        foreach (var entry in entries) dictionary.Items.Add(entry.Key, entry.Value);
        return dictionary;
    }

    private static PdfArray ArrayOf(params PdfObject[] values) {
        var array = new PdfArray();
        array.Items.AddRange(values);
        return array;
    }

    private static PdfArray Numbers(params double[] values) => ArrayOf(values.Select(value => (PdfObject)new PdfNumber(value)).ToArray());

    private static PdfName Name(string value) => new(value);
}
