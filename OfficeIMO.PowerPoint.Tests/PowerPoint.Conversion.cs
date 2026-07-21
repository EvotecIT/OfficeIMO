using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointConversionTests {
    [Fact]
    public void Convert_PptxToPptAndBack_UsesUnifiedCompatibilityReport() {
        using var files = new ConversionFiles();
        string pptx = files.Path("source.pptx");
        string ppt = files.Path("converted.ppt");
        string roundTrip = files.Path("roundtrip.pptx");
        CreateTextPresentation(pptx);

        PowerPointPresentationConversionResult toPpt = PowerPointPresentation.Convert(pptx, ppt);

        Assert.Equal(ppt, toPpt.RequireNoLoss());
        Assert.Equal("PowerPoint.Pptx", toPpt.Report.SourceFormatDescriptor.Id);
        Assert.Equal("PowerPoint.Ppt", toPpt.Report.DestinationFormatDescriptor.Id);
        Assert.True(toPpt.Report.Compatibility.IsStrictlyCompatible);

        PowerPointPresentationConversionResult toPptx = PowerPointPresentation.Convert(ppt, roundTrip);

        Assert.Equal(roundTrip, toPptx.RequireNoLoss());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(roundTrip);
        Assert.Contains(reopened.Slides[0].TextBoxes,
            textBox => textBox.Text.Contains("Conversion contract", StringComparison.Ordinal));
    }

    [Fact]
    public void Convert_AllowsModernDocumentKindChangeDespiteSharedBroadFormat() {
        using var files = new ConversionFiles();
        string pptx = files.Path("source.pptx");
        string potx = files.Path("template.potx");
        CreateTextPresentation(pptx);

        PowerPointPresentationConversionResult result = PowerPointPresentation.Convert(pptx, potx);

        Assert.Equal(PowerPointFileFormat.Pptx, result.Report.SourceFormat);
        Assert.Equal(PowerPointFileFormat.Pptx, result.Report.DestinationFormat);
        Assert.Equal(OfficeDocumentKind.Document, result.Report.SourceFormatDescriptor.DocumentKind);
        Assert.Equal(OfficeDocumentKind.Template, result.Report.DestinationFormatDescriptor.DocumentKind);
        Assert.True(File.Exists(potx));
        using PowerPointPresentation reopened = PowerPointPresentation.Load(potx);
        Assert.Equal("PowerPoint.Potx", reopened.SourceFormatDescriptor.Id);
    }

    [Fact]
    public void Convert_PreferVisualReportsChartRasterization() {
        using var files = new ConversionFiles();
        string source = files.Path("chart.pptx");
        string destination = files.Path("chart.ppt");
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(source)) {
            presentation.AddSlide(SlideLayoutValues.Blank).AddChart();
            presentation.Save();
        }

        PowerPointPresentationConversionException blocked = Assert.Throws<PowerPointPresentationConversionException>(() =>
            PowerPointPresentation.Convert(source, destination));
        Assert.Equal(PowerPointPresentationConversionFailureReason.DataLossBlocked, blocked.Reason);
        Assert.Contains(blocked.Result.Report.Compatibility.Findings,
            finding => finding.State == OfficeCompatibilityState.Blocked
                && finding.Code == "PPT-WRITE-CHART-CONVERTED");

        PowerPointPresentationConversionResult converted = PowerPointPresentation.Convert(
            source,
            destination,
            new PowerPointPresentationConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.PreferVisual
            });

        Assert.Contains(converted.Report.Compatibility.Findings,
            finding => finding.State == OfficeCompatibilityState.Rasterized
                && finding.Code == "PPT-WRITE-CHART-CONVERTED"
                && (finding.Impact & OfficeCompatibilityImpact.Editability) != 0);
        Assert.True(converted.HasLoss);
        using PowerPointPresentation reopened = PowerPointPresentation.Load(destination);
        Assert.Empty(reopened.Slides[0].Charts);
        Assert.Single(reopened.Slides[0].Pictures);
    }

    [Fact]
    public void Convert_PreferEditableBlocksStaticChartFallback() {
        using var files = new ConversionFiles();
        string source = files.Path("chart.pptx");
        string destination = files.Path("chart.ppt");
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(source)) {
            presentation.AddSlide(SlideLayoutValues.Blank).AddChart();
            presentation.Save();
        }

        PowerPointPresentationConversionException exception = Assert.Throws<PowerPointPresentationConversionException>(() =>
            PowerPointPresentation.Convert(
                source,
                destination,
                new PowerPointPresentationConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.PreferEditable
                }));

        Assert.Equal(PowerPointPresentationConversionFailureReason.DataLossBlocked, exception.Reason);
        Assert.Contains(exception.Result.Report.Compatibility.Findings,
            finding => finding.Code == "PPT-WRITE-CHART-CONVERTED"
                && finding.State == OfficeCompatibilityState.Blocked);
        Assert.False(File.Exists(destination));
    }

    [Fact]
    public void Convert_EditableAndVisualModesBlockGenericUnmappedContent() {
        using var files = new ConversionFiles();
        string source = files.Path("custom-xml.pptx");
        CreateTextPresentation(source);
        using (PresentationDocument package = PresentationDocument.Open(source, true)) {
            CustomXmlPart customXml = package.PresentationPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var data = new MemoryStream(Encoding.UTF8.GetBytes("<compatibility>unmapped</compatibility>"));
            customXml.FeedData(data);
        }

        foreach (OfficeCompatibilityMode mode in new[] {
                     OfficeCompatibilityMode.PreferEditable,
                     OfficeCompatibilityMode.PreferVisual
                 }) {
            string destination = files.Path("custom-xml-" + mode + ".ppt");
            var options = new PowerPointPresentationConversionOptions { CompatibilityMode = mode };

            PowerPointPresentationConversionReport analysis = PowerPointPresentation.AnalyzeConversion(
                source,
                destination,
                options);

            Assert.Contains(analysis.Compatibility.Findings,
                finding => finding.Code == "PPT-WRITE-CUSTOM-XML"
                    && finding.State == OfficeCompatibilityState.Blocked
                    && finding.RepresentsLoss);
            PowerPointPresentationConversionException blocked = Assert.Throws<PowerPointPresentationConversionException>(() =>
                PowerPointPresentation.Convert(source, destination, options));
            Assert.Equal(PowerPointPresentationConversionFailureReason.DataLossBlocked, blocked.Reason);
            Assert.False(File.Exists(destination));
        }
    }

    [Fact]
    public void Convert_PreservationOnlyRetainsSourceInLegacyCarrier() {
        using var files = new ConversionFiles();
        string source = files.Path("chart.pptx");
        string destination = files.Path("chart.ppt");
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(source)) {
            presentation.AddSlide(SlideLayoutValues.Blank).AddChart();
            presentation.Save();
        }
        byte[] sourceBytes = File.ReadAllBytes(source);

        PowerPointPresentationConversionResult result = PowerPointPresentation.Convert(
            source,
            destination,
            new PowerPointPresentationConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.PreservationOnly
            });

        Assert.Contains(result.Report.Compatibility.Findings,
            finding => finding.Code == "PPT-WRITE-CHART-CONVERTED"
                && finding.State == OfficeCompatibilityState.Rasterized);
        Assert.Contains(result.Report.Compatibility.Findings,
            finding => finding.Code == "PowerPoint.SourceCarrier.Embedded"
                && finding.State == OfficeCompatibilityState.EmbeddedSource);
        using PowerPointPresentation reopened = PowerPointPresentation.Load(destination);
        Assert.Empty(reopened.Slides[0].Charts);
        Assert.Single(reopened.Slides[0].Pictures);
        Assert.True(reopened.TryGetCompatibilitySourcePayload(out OfficeCompatibilitySourcePayload? payload, out string? error), error);
        Assert.NotNull(payload);
        Assert.Equal("PowerPoint.Pptx", payload!.FormatId);
        Assert.Equal(OfficeCompatibilityMode.PreservationOnly, payload.Mode);
        Assert.Equal(sourceBytes, payload.ToArray());
    }

    [Fact]
    public void Convert_LegacyToModernPreservationOnlyRetainsImmediateSource() {
        using var files = new ConversionFiles();
        string source = files.Path("source.pptx");
        string legacy = files.Path("source.ppt");
        string destination = files.Path("restored.pptx");
        CreateTextPresentation(source);
        PowerPointPresentation.Convert(source, legacy).RequireNoLoss();
        byte[] legacyBytes = File.ReadAllBytes(legacy);

        PowerPointPresentationConversionResult result = PowerPointPresentation.Convert(
            legacy,
            destination,
            new PowerPointPresentationConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.PreservationOnly
            });

        Assert.Contains(result.Report.Compatibility.Findings,
            finding => finding.Code == "PowerPoint.SourceCarrier.Embedded"
                && finding.State == OfficeCompatibilityState.EmbeddedSource);
        using PowerPointPresentation reopened = PowerPointPresentation.Load(destination);
        Assert.True(reopened.TryGetCompatibilitySourcePayload(out OfficeCompatibilitySourcePayload? payload, out string? error), error);
        Assert.NotNull(payload);
        Assert.Equal("PowerPoint.Ppt", payload!.FormatId);
        Assert.Equal(legacyBytes, payload.ToArray());
    }

    [Fact]
    public void Convert_ClassifiedLegacyAddInDestination_IsExplicitlyBlocked() {
        using var files = new ConversionFiles();
        string source = files.Path("source.pptx");
        string destination = files.Path("add-in.ppa");
        CreateTextPresentation(source);

        PowerPointPresentationConversionException exception = Assert.Throws<PowerPointPresentationConversionException>(() =>
            PowerPointPresentation.Convert(source, destination));

        Assert.Equal(PowerPointPresentationConversionFailureReason.DestinationFeatureUnsupported, exception.Reason);
        Assert.Contains(exception.Result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "PowerPoint.LegacyDestination.NotWritable"
                && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked);
        Assert.False(File.Exists(destination));
    }

    [Fact]
    public void AnalyzeConversion_ReportsVisualFallbackWithoutCreatingOutput() {
        using var files = new ConversionFiles();
        string source = files.Path("chart.pptx");
        string destination = files.Path("chart.ppt");
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(source)) {
            presentation.AddSlide(SlideLayoutValues.Blank).AddChart();
            presentation.Save();
        }

        PowerPointPresentationConversionReport report = PowerPointPresentation.AnalyzeConversion(
            source,
            destination,
            new PowerPointPresentationConversionOptions { CompatibilityMode = OfficeCompatibilityMode.PreferVisual });

        Assert.Contains(report.Compatibility.Findings,
            finding => finding.State == OfficeCompatibilityState.Rasterized);
        Assert.False(File.Exists(destination));
    }

    [Fact]
    public void Convert_DecryptedLegacySourceReportsPasswordProtectionLoss() {
        using var files = new ConversionFiles();
        const string password = "openpass";
        string source = files.Path("encrypted.ppt");
        string blockedDestination = files.Path("blocked.pptx");
        string allowedDestination = files.Path("allowed.pptx");
        using (PowerPointPresentation presentation = PowerPointPresentation.Create()) {
            presentation.AddSlide(SlideLayoutValues.Blank).AddTextBox("Encrypted conversion source");
            File.WriteAllBytes(source, presentation.ToEncryptedBytes(password, PowerPointFileFormat.Ppt));
        }

        var blockedOptions = new PowerPointPresentationConversionOptions {
            LoadOptions = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions { Password = password }
            }
        };
        PowerPointPresentationConversionReport analysis = PowerPointPresentation.AnalyzeConversion(
            source,
            blockedDestination,
            blockedOptions);

        Assert.True(analysis.Compatibility.HasSecurityImpact);
        Assert.Contains(analysis.Compatibility.Findings,
            finding => finding.Code == "PPT-ENCRYPTION-DECRYPTED"
                && finding.State == OfficeCompatibilityState.Blocked
                && finding.RepresentsLoss);
        PowerPointPresentationConversionException blocked = Assert.Throws<PowerPointPresentationConversionException>(() =>
            PowerPointPresentation.Convert(source, blockedDestination, blockedOptions));
        Assert.Equal(PowerPointPresentationConversionFailureReason.DataLossBlocked, blocked.Reason);
        Assert.False(File.Exists(blockedDestination));

        PowerPointPresentationConversionResult allowed = PowerPointPresentation.Convert(
            source,
            allowedDestination,
            new PowerPointPresentationConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                LoadOptions = new PowerPointLoadOptions {
                    LegacyPptImportOptions = new LegacyPptImportOptions { Password = password }
                }
            });
        Assert.True(allowed.Report.Compatibility.HasSecurityImpact);
        Assert.Contains(allowed.Report.Compatibility.Findings,
            finding => finding.Code == "PPT-ENCRYPTION-DECRYPTED"
                && finding.State == OfficeCompatibilityState.Dropped
                && finding.RepresentsLoss);
        Assert.True(File.Exists(allowedDestination));
    }

    [Fact]
    public void Convert_SignedModernSourceIsIncludedInAnalysisAndStructuredFailure() {
        using var files = new ConversionFiles();
        string source = files.Path("signed.pptx");
        string blockedDestination = files.Path("blocked.potx");
        string allowedDestination = files.Path("allowed.potx");
        CreateTextPresentation(source);
        AddSyntheticSignature(source);

        PowerPointPresentationConversionReport analysis = PowerPointPresentation.AnalyzeConversion(
            source,
            blockedDestination);

        Assert.True(analysis.Compatibility.HasSecurityImpact);
        Assert.Contains(analysis.Compatibility.Findings,
            finding => finding.Code == "PowerPoint.DigitalSignature.Invalidated"
                && finding.State == OfficeCompatibilityState.Blocked
                && finding.RepresentsLoss);
        PowerPointPresentationConversionException blocked = Assert.Throws<PowerPointPresentationConversionException>(() =>
            PowerPointPresentation.Convert(source, blockedDestination));
        Assert.Equal(PowerPointPresentationConversionFailureReason.DataLossBlocked, blocked.Reason);
        Assert.False(File.Exists(blockedDestination));

        PowerPointPresentationConversionResult allowed = PowerPointPresentation.Convert(
            source,
            allowedDestination,
            new PowerPointPresentationConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                SignatureMutationPolicy = PowerPointSignatureMutationPolicy.RemoveInvalidatedSignatures
            });
        Assert.True(allowed.Report.Compatibility.HasSecurityImpact);
        Assert.Contains(allowed.Report.Compatibility.Findings,
            finding => finding.Code == "PowerPoint.DigitalSignature.Invalidated"
                && finding.State == OfficeCompatibilityState.Dropped
                && finding.RepresentsLoss);
        using PresentationDocument package = PresentationDocument.Open(allowedDestination, false);
        Assert.Null(package.DigitalSignatureOriginPart);
    }

    private static void CreateTextPresentation(string path) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
        presentation.AddSlide(SlideLayoutValues.Blank).AddTextBox("Conversion contract");
        presentation.Save();
    }

    private static void AddSyntheticSignature(string path) {
        using PresentationDocument document = PresentationDocument.Open(path, true);
        DigitalSignatureOriginPart origin = document.AddDigitalSignatureOriginPart();
        XmlSignaturePart signature = origin.AddNewPart<XmlSignaturePart>();
        using var data = new MemoryStream(Encoding.UTF8.GetBytes(
            "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo/><SignatureValue>AA==</SignatureValue></Signature>"));
        signature.FeedData(data);
    }

    private sealed class ConversionFiles : IDisposable {
        private readonly string _directory = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            "OfficeIMO-PowerPoint-Conversion-" + Guid.NewGuid().ToString("N"));

        internal ConversionFiles() => Directory.CreateDirectory(_directory);

        internal string Path(string fileName) => System.IO.Path.Combine(_directory, fileName);

        public void Dispose() {
            try {
                if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
            } catch {
            }
        }
    }
}
