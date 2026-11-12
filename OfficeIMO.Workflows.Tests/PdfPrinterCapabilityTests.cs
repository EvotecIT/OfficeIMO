using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPrinterCapabilityTests {
    // Captured from `lpoptions -p <queue> -l` and `lpoptions -p <queue>` for an HP LaserJet Pro M404 IPP queue on macOS.
    private const string LaserJetChoices = """
        PageSize/Media Size: 100x150mm 184x260mm 195x270mm 4x6 5x8 *A4 A5 A6 B5 B6 DoublePostcardRotated Env10 EnvC5 EnvDL EnvMonarch Executive FanFoldGermanLegal ISOB5 Legal Letter Oficio Postcard Statement roc16k Custom.WIDTHxHEIGHT
        InputSlot/Media Source: Tray1 Tray2 *Auto
        MediaType/Media Type: *Stationery Com.hp.stationeryLightweight Labels Envelope
        ColorModel/Output Mode: *RGB Gray DeviceGray DeviceRGB AdobeRGB
        Duplex/Duplex: *None DuplexNoTumble DuplexTumble
        cupsPrintQuality/cupsPrintQuality: Draft *Normal High
        """;
    private const string LaserJetDestination = "device-uri=ipp://Drukarka%20HP._ipp._tcp.local./ printer-info='Drukarka HP' printer-location printer-make-and-model='HP LaserJet Pro M404-M405' printer-type=16781332";

    [Fact]
    public void CupsQueueReportsDefaultMediaSizesTraysAndDeviceTypeBits() {
        PdfPrinterCapabilities capabilities = CupsPdfPrinter.ParseCapabilities("Drukarka_HP", LaserJetChoices, LaserJetDestination);
        Assert.Equal("A4", capabilities.DefaultMedia!.Id);
        Assert.Equal(595.28, capabilities.DefaultMedia.Size!.Value.Width.ToPoints(), 1);
        Assert.DoesNotContain(capabilities.Media, media => media.Id.StartsWith("Custom.", StringComparison.Ordinal));
        Assert.Equal(283.46, capabilities.Media.Single(media => media.Id == "100x150mm").Size!.Value.Width.ToPoints(), 1);
        Assert.Equal(432, capabilities.Media.Single(media => media.Id == "4x6").Size!.Value.Height.ToPoints(), 1);
        Assert.Equal(515.91, capabilities.Media.Single(media => media.Id == "B5").Size!.Value.Width.ToPoints(), 1); // PPD B5 is JIS.
        Assert.Null(capabilities.Media.Single(media => media.Id == "Oficio").Size);
        Assert.Equal(new[] { "InputSlot=Tray1", "InputSlot=Tray2", "InputSlot=Auto" }, capabilities.PaperSources.Select(source => source.Id));
        // printer-type marks this monochrome duplex printer even though the PPD offers RGB output modes.
        Assert.True(capabilities.SupportsDuplex);
        Assert.False(capabilities.SupportsColor);
    }

    [Fact]
    public void CupsIppEverywhereNamesAndResolutionsParseWithoutPrinterType() {
        PdfPrinterCapabilities capabilities = CupsPdfPrinter.ParseCapabilities("queue",
            "media/Media: *iso_a4_210x297mm na_letter_8.5x11in\nsides/2-Sided: *one-sided two-sided-long-edge\n" +
            "print-color-mode/Color: monochrome *color\nResolution/Output Resolution: 300dpi *600dpi 1200x600dpi bad\nmedia-source/Source: *auto main",
            "printer-info=queue");
        Assert.Equal("iso_a4_210x297mm", capabilities.DefaultMedia!.Id);
        Assert.Equal(612, capabilities.Media[1].Size!.Value.Width.ToPoints(), 1);
        Assert.True(capabilities.SupportsDuplex);
        Assert.True(capabilities.SupportsColor);
        Assert.Equal(new PdfPrinterResolution[] { new(300, 300), new(600, 600), new(1200, 600) }, capabilities.Resolutions);
        Assert.Equal(new[] { "media-source=auto", "media-source=main" }, capabilities.PaperSources.Select(source => source.Id));
    }

    [Fact]
    public void CupsChoicesRejectOptionFragmentsAndUnreportedFeaturesStayUnknown() {
        PdfPrinterCapabilities capabilities = CupsPdfPrinter.ParseCapabilities("queue", "InputSlot/Source: Tray1 Tray1 bad=value /evil", string.Empty);
        Assert.Equal(new[] { "InputSlot=Tray1" }, capabilities.PaperSources.Select(source => source.Id));
        Assert.Empty(capabilities.Media);
        Assert.Null(capabilities.DefaultMedia);
        Assert.Null(capabilities.SupportsDuplex);
        Assert.Null(capabilities.SupportsColor);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("tray\n-o")]
    public void DeliveryRejectsMalformedPaperSource(string source) {
        var options = new PdfPrintDeliveryOptions { PrinterName = "queue", PaperSourceId = source };
        Assert.Throws<ArgumentException>(() => options.Snapshot());
    }

    [Fact]
    public void WindowsDriverListsBecomeSizedMediaWithTheDriverDefault() {
        PdfPrinterCapabilities capabilities = WindowsPrinterSettings.BuildCapabilities("Office", new WindowsPrinterSettings.DriverCapabilities {
            PaperCodes = [1, 9, 300], PaperSizes = [(2159, 2794), (2100, 2970), (0, 0)], PaperNames = ["Letter", "A4", ""],
            DefaultPaperCode = 9, PaperSources = [new("7", "Auto")], Duplex = 1, ColorDevice = 0, Collate = 1, Copies = 999,
            Resolutions = [(600, 600), (1200, 1200), (-1, -1)]
        });
        Assert.Equal("9", capabilities.DefaultMedia!.Id);
        Assert.Equal(595.28, capabilities.DefaultMedia.Size!.Value.Width.ToPoints(), 1);
        Assert.Equal("300", capabilities.Media[2].Name);
        Assert.Null(capabilities.Media[2].Size);
        Assert.True(capabilities.SupportsDuplex);
        Assert.False(capabilities.SupportsColor);
        Assert.True(capabilities.SupportsCollation);
        Assert.Equal(999, capabilities.MaximumCopies);
        Assert.Equal(2, capabilities.Resolutions.Count);
        Assert.Null(WindowsPrinterSettings.BuildCapabilities("Office", new WindowsPrinterSettings.DriverCapabilities()).SupportsDuplex);
        Assert.Throws<InvalidOperationException>(() => WindowsPrinterSettings.BuildCapabilities("Office",
            new WindowsPrinterSettings.DriverCapabilities { PaperCodes = [1, 9], PaperSizes = [(2159, 2794)], PaperNames = ["Letter", "A4"] }));
    }

    [Fact]
    public void WindowsPaperCodesStayUsableWhenTheDriverOmitsNamesOrSizes() {
        PdfPrinterCapabilities unnamed = WindowsPrinterSettings.BuildCapabilities("Office", new WindowsPrinterSettings.DriverCapabilities {
            PaperCodes = [1, 9], PaperSizes = [(2159, 2794), (2100, 2970)], DefaultPaperCode = 9
        });
        Assert.Equal(["1", "9"], unnamed.Media.Select(media => media.Name));
        Assert.Equal(595.28, unnamed.DefaultMedia!.Size!.Value.Width.ToPoints(), 1);
        PdfPrinterCapabilities unsized = WindowsPrinterSettings.BuildCapabilities("Office", new WindowsPrinterSettings.DriverCapabilities {
            PaperCodes = [1, 9], PaperNames = ["Letter", "A4"]
        });
        Assert.Equal(["Letter", "A4"], unsized.Media.Select(media => media.Name));
        Assert.All(unsized.Media, media => Assert.Null(media.Size));
        Assert.Equal(9, WindowsPrinterSettings.ResolvePaperCode(PdfPageSizes.A4, unsized.Media));
    }

    [Fact]
    public void WindowsPaperCodesPreferTheDriverThenStandardCodesThenCustomSize() {
        IReadOnlyList<PdfPrinterMedia> driver = [new("120", "Vendor A4", PdfPageSizes.A4), new("256", "Custom", PdfPageSizes.A3)];
        Assert.Equal(120, WindowsPrinterSettings.ResolvePaperCode(PdfPageSizes.A4.Landscape(), driver));
        Assert.Equal(8, WindowsPrinterSettings.ResolvePaperCode(PdfPageSizes.A3, driver));
        Assert.Equal(1, WindowsPrinterSettings.ResolvePaperCode(PdfPageSizes.Letter, []));
        Assert.Equal(256, WindowsPrinterSettings.ResolvePaperCode(new PdfPageSize(OfficeLength.Points(250), OfficeLength.Points(400)), []));
    }

    [Fact]
    public void DevModeCompositionWritesDriverCopiesCustomPaperSourceAndDuplex() {
        byte[] mode = DriverDefaults(fields: WindowsPrinterSettings.FieldPaperLength | 0x40000);
        var paper = new PdfPageSize(OfficeLength.Points(400), OfficeLength.Points(250));
        var request = new WindowsPrinterSettings.DevModeRequest(paper, WindowsPrinterSettings.CustomPaperCode, PdfPrintDuplex.LongEdge, "260", 3);
        WindowsPrinterSettings.ComposeDevMode(mode, request);
        Assert.Equal(2, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModeOrientation));
        Assert.Equal(256, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModePaperSize));
        Assert.Equal(1411, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModePaperLength)); // 400 pt in tenths of a millimetre.
        Assert.Equal(882, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModePaperWidth));
        Assert.Equal(3, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModeCopies));
        Assert.Equal(1, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModeCollate));
        Assert.Equal(260, BitConverter.ToUInt16(mode, WindowsPrinterSettings.DevModeDefaultSource));
        Assert.Equal(2, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModeDuplex));
        int fields = BitConverter.ToInt32(mode, WindowsPrinterSettings.DevModeFields);
        Assert.Equal(0x40000, fields & 0x40000); // Unrelated driver fields survive.
        foreach (int field in new[] { WindowsPrinterSettings.FieldCopies, WindowsPrinterSettings.FieldCollate, WindowsPrinterSettings.FieldDefaultSource,
            WindowsPrinterSettings.FieldDuplex, WindowsPrinterSettings.FieldPaperLength, WindowsPrinterSettings.FieldPaperWidth })
            Assert.Equal(field, fields & field);
        Assert.Null(WindowsPrinterSettings.FindRejectedSetting(mode, request));
        Assert.Equal(0xEE, mode[^1]); // Driver-private trailing data is preserved.

        var standard = new WindowsPrinterSettings.DevModeRequest(PdfPageSizes.A4, 9, PdfPrintDuplex.PrinterDefault, null, 1);
        WindowsPrinterSettings.ComposeDevMode(mode, standard);
        Assert.Equal(0, BitConverter.ToInt32(mode, WindowsPrinterSettings.DevModeFields) & (WindowsPrinterSettings.FieldPaperLength | WindowsPrinterSettings.FieldPaperWidth));
        Assert.Equal(1, BitConverter.ToInt16(mode, WindowsPrinterSettings.DevModeCopies));
    }

    [Fact]
    public void DevModeVerificationNamesTheSettingTheDriverChanged() {
        byte[] mode = DriverDefaults(fields: 0);
        var request = new WindowsPrinterSettings.DevModeRequest(PdfPageSizes.A4, 9, PdfPrintDuplex.ShortEdge, "7", 2);
        WindowsPrinterSettings.ComposeDevMode(mode, request);
        byte[] copies = (byte[])mode.Clone();
        copies[WindowsPrinterSettings.DevModeCopies] = 1;
        Assert.Equal("copies", WindowsPrinterSettings.FindRejectedSetting(copies, request));
        byte[] collate = (byte[])mode.Clone();
        collate[WindowsPrinterSettings.DevModeCollate] = 0;
        Assert.Equal("collation", WindowsPrinterSettings.FindRejectedSetting(collate, request));
        byte[] duplex = (byte[])mode.Clone();
        duplex[WindowsPrinterSettings.DevModeDuplex] = 1;
        Assert.Equal("duplex", WindowsPrinterSettings.FindRejectedSetting(duplex, request));
        byte[] source = (byte[])mode.Clone();
        source[WindowsPrinterSettings.DevModeDefaultSource] = 8;
        Assert.Equal("paper source", WindowsPrinterSettings.FindRejectedSetting(source, request));
        Assert.Throws<InvalidOperationException>(() => WindowsPrinterSettings.ComposeDevMode(new byte[100], request));
    }

    [Theory]
    [InlineData(3, true, 999, true, 3)]
    [InlineData(3, true, 2, true, 1)]
    [InlineData(3, false, 999, true, 1)]
    [InlineData(3, true, 999, false, 1)]
    [InlineData(1, true, 999, true, 1)]
    public void DriverCopiesRequireCollationCapacityAndOnePaper(int requested, bool collates, int maximum, bool singlePaper, short expected) {
        var capabilities = new PdfPrinterCapabilities { PrinterName = "Office", SupportsCollation = collates, MaximumCopies = maximum };
        Assert.Equal(expected, WindowsPrinterSettings.DriverCopies(requested, capabilities, singlePaper));
    }

    private static byte[] DriverDefaults(int fields) {
        byte[] mode = new byte[220 + 16];
        BitConverter.GetBytes((ushort)220).CopyTo(mode, WindowsPrinterSettings.DevModeSize);
        BitConverter.GetBytes(fields).CopyTo(mode, WindowsPrinterSettings.DevModeFields);
        for (int index = 220; index < mode.Length; index++) mode[index] = 0xEE;
        return mode;
    }
}
