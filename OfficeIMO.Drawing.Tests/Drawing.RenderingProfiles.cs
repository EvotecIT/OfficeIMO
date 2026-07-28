using System;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRenderingProfileTests {
    [Fact]
    public void ManagedProfileReplacesFormatNeutralRenderingSettings() {
        var originalProvider = new DecliningTextShapingProvider();
        var options = new OfficeImageExportOptions {
            TextShapingProvider = originalProvider,
            TextShapingLanguage = "pl",
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };

        OfficeImageExportOptions returned = options.UseRenderingProfile(OfficeRenderingProfile.Managed);

        Assert.Same(options, returned);
        Assert.Same(OfficeManagedTextShapingProvider.Instance, options.TextShapingProvider);
        Assert.Null(options.TextShapingLanguage);
        Assert.False(options.Policy.RequireNoLoss);
        Assert.NotSame(OfficeRenderingProfile.Managed.Fonts, options.Fonts);
    }

    [Fact]
    public void OverlayPreservesExistingOptionalServicesWhenProfileDoesNotProvideThem() {
        var provider = new DecliningTextShapingProvider();
        var codec = new DecliningImageCodec();
        var options = new OfficeImageExportOptions {
            TextShapingProvider = provider,
            TextShapingLanguage = "pl",
            ImageCodec = codec
        };
        var profile = new OfficeRenderingProfile(
            "strict",
            policy: new OfficeImageExportPolicy { RequireNoOmissions = true });

        options.UseRenderingProfile(profile, OfficeRenderingProfileApplyMode.Overlay);

        Assert.Same(provider, options.TextShapingProvider);
        Assert.Equal("pl", options.TextShapingLanguage);
        Assert.Same(codec, options.ImageCodec);
        Assert.True(options.Policy.RequireNoOmissions);
    }

    [Fact]
    public void ProfileOwnsDefensiveFontAndPolicySnapshots() {
        var policy = new OfficeImageExportPolicy { RequireNoLoss = true };
        var profile = new OfficeRenderingProfile(" deterministic ", policy: policy);

        policy.RequireNoLoss = false;
        OfficeImageExportPolicy first = profile.Policy;
        first.RequireNoLoss = false;

        Assert.Equal("deterministic", profile.Name);
        Assert.True(profile.Policy.RequireNoLoss);
        Assert.NotSame(profile.Fonts, profile.Fonts);
    }

    [Fact]
    public void RejectsUnknownApplicationMode() {
        var options = new OfficeImageExportOptions();

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            options.UseRenderingProfile(OfficeRenderingProfile.Managed, (OfficeRenderingProfileApplyMode)42));
    }

    [Fact]
    public void FluentApplicationPreservesDerivedOptionType() {
        ExcelImageExportOptions options = new ExcelImageExportOptions()
            .UseRenderingProfile(OfficeRenderingProfile.Managed);

        Assert.Same(OfficeManagedTextShapingProvider.Instance, options.TextShapingProvider);
    }

    private sealed class DecliningTextShapingProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => null;
    }

    private sealed class DecliningImageCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = null;
            return false;
        }
    }
}
