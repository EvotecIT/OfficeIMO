using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveCompoundAdmissionCompatibilityTests {
    [Theory]
    [InlineData("full")]
    [InlineData("selective")]
    [InlineData("view")]
    [InlineData("directory")]
    [InlineData("detector")]
    public void ExistingNativeTryContractsRetainDefaultQuotaFailureMessages(string route) {
        byte[] bytes = ArchiveFixtures.Msi();
        using var source = new MemoryStream(bytes);
        var options = new OfficeCompoundReadOptions(maxDirectoryEntries: 3);
        string? error;
        bool result;
        switch (route) {
            case "full": result = OfficeCompoundFileReader.TryRead(bytes, options, out OfficeCompoundFile? _, out error); break;
            case "selective": result = OfficeCompoundFileReader.TryReadSelective(source, options, (_, _) => false,
                (_, _) => throw new InvalidOperationException("No destination is needed."), out _, out error); break;
            case "view": result = OfficeCompoundFileReader.TryOpenStream(source, options, (_, _) => true, true, default, out _, out error); break;
            case "directory": result = OfficeCompoundFileReader.TryInspectDirectory(source, long.MaxValue, 3, out _, out error); break;
            default: result = OfficeCompoundDocumentDetector.Detect(source, long.MaxValue, 3, out error)
                != OfficeCompoundDocumentDetector.DocumentKind.UnknownCompound; break;
        }
        Assert.False(result); Assert.Contains("entry count exceeds 3", error);
        Assert.True(source.CanRead);
        if (route is "selective" or "view" or "detector") Assert.Equal(0, source.Position);
    }

    [Theory]
    [InlineData("directory")]
    [InlineData("stream")]
    [InlineData("count")]
    [InlineData("input")]
    public void DefaultSessionRetainsItsInvalidDataFailureAndCallerOwnership(string bound) {
        using var source = new MemoryStream(ArchiveFixtures.Msi());
        var options = bound switch {
            "directory" => new OfficeCompoundReadOptions(maxDirectoryEntries: 3),
            "stream" => new OfficeCompoundReadOptions(maxStreamBytes: 1),
            "count" => new OfficeCompoundReadOptions(maxStreamCount: 1),
            _ => new OfficeCompoundReadOptions(maxTotalStreamBytes: 1)
        };
        Assert.Throws<InvalidDataException>(() => new OfficeCompoundFileReader.ReadSession(source,
            options, leaveOpen: true, default));
        Assert.Equal(0, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("stream", "has unsupported size")]
    [InlineData("count", "stream count")]
    [InlineData("allocation", "allocation table counts exceed configured or physical bounds")]
    public void DefaultFullReaderPreservesItsExistingByteCountAndAllocationErrors(string bound, string message) {
        var options = bound switch {
            "stream" => new OfficeCompoundReadOptions(maxStreamBytes: 1),
            "count" => new OfficeCompoundReadOptions(maxStreamCount: 1),
            _ => new OfficeCompoundReadOptions(maxTotalStreamBytes: 1)
        };
        Assert.False(OfficeCompoundFileReader.TryRead(ArchiveFixtures.Msi(), options, out OfficeCompoundFile? _, out string? error));
        Assert.Contains(message, error);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ExplicitDetectorAndContainsOptInRetainsQuotaAndCursorIdentity(bool tokenOverload, bool contains) {
        using var source = new MemoryStream(ArchiveFixtures.Msi());
        OfficeCompoundStreamLimitExceededException error = Assert.Throws<OfficeCompoundStreamLimitExceededException>(() => {
            if (contains) {
                if (tokenOverload) OfficeCompoundFileReader.TryContainsStreamPath(source, "Workbook", long.MaxValue, 3,
                    default, out _, out _, preserveResourceLimits: true);
                else OfficeCompoundFileReader.TryContainsStreamPath(source, "Workbook", long.MaxValue, 3,
                    out _, out _, preserveResourceLimits: true);
            } else {
                if (tokenOverload) OfficeCompoundDocumentDetector.Detect(source, long.MaxValue, 3, default, out _, preserveResourceLimits: true);
                else OfficeCompoundDocumentDetector.Detect(source, long.MaxValue, 3, out _, preserveResourceLimits: true);
            }
        });
        Assert.Equal("MaxDirectoryEntries", error.LimitName); Assert.Equal(3, error.MaximumValue);
        Assert.True(source.CanRead);
        if (!contains) Assert.Equal(0, source.Position);
    }

    [Fact]
    public void DefaultDirectoryInputLimitStillReturnsItsOriginalError() {
        using var source = new MemoryStream(ArchiveFixtures.Msi());
        Assert.False(OfficeCompoundFileReader.TryInspectDirectory(source, 1, 65536, out _, out string? error));
        Assert.Equal("The compound file exceeds the configured input bound of 1.", error);
    }

    [Theory]
    [InlineData("full")]
    [InlineData("selective")]
    [InlineData("view")]
    public void ExistingCallerValidatorQuotaIsNeverConvertedToAMalformedResult(string route) {
        byte[] bytes = ArchiveFixtures.Msi();
        using var source = new MemoryStream(bytes);
        var original = new OfficeCompoundStreamLimitExceededException("CallerLimit", 2, 1);
        var options = new OfficeCompoundReadOptions(streamSizeValidator: (_, _) => throw original);
        Exception error = Assert.Throws<OfficeCompoundStreamLimitExceededException>(() => {
            switch (route) {
                case "full": OfficeCompoundFileReader.TryRead(bytes, options, out OfficeCompoundFile? _, out _); break;
                case "selective": OfficeCompoundFileReader.TryReadSelective(source, options, (_, _) => false,
                    (_, _) => throw new InvalidOperationException("No destination is needed."), out _, out _); break;
                default: OfficeCompoundFileReader.TryOpenStream(source, options, (_, _) => true, true, default, out _, out _); break;
            }
        });
        Assert.Same(original, error); Assert.True(source.CanRead);
        if (route is "selective" or "view") Assert.Equal(0, source.Position);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PublicPackageSecurityKeepsItsPartCountReportInsteadOfExposingAnInternalException(bool stream) {
        byte[] bytes = ArchiveFixtures.Msi();
        var options = new OfficePackageSecurityOptions { MaxPartCount = 1 };
        using var source = new MemoryStream(bytes); source.Position = 9;
        OfficePackageSecurityReport report = stream ? OfficePackageSecurityInspector.Inspect(source, options)
            : OfficePackageSecurityInspector.Inspect(bytes, options);
        var finding = Assert.Single(report.Findings, item => item.Rule == OfficePackageSecurityRule.PartCount);
        Assert.Equal(1, finding.Limit); Assert.Equal(2, finding.ObservedValue);
        Assert.DoesNotContain(report.Findings, item => item.Rule == OfficePackageSecurityRule.MalformedPackage);
        Assert.Equal(9, source.Position); Assert.True(source.CanRead);
    }
}
