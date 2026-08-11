using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class OfficeColorSpaceConverterTests {
    [Fact]
    public void CalibratedRgb_AppliesGammaAndColumnMajorMatrix() {
        OfficeColor color = OfficeColorSpaceConverter.FromCalibratedRgb(
            0.5D,
            0.25D,
            0.75D,
            0.95047D,
            1D,
            1.08883D,
            gamma: new[] { 2D, 2D, 2D },
            matrix: new[] {
                0.4124564D, 0.2126729D, 0.0193339D,
                0.3575761D, 0.7151522D, 0.119192D,
                0.1804375D, 0.072175D, 0.9503041D
            });

        Assert.InRange(color.R, 136, 138);
        Assert.InRange(color.G, 70, 72);
        Assert.InRange(color.B, 197, 199);
    }

    [Fact]
    public void LabAndCmyk_ProduceStableSrgbPrimaries() {
        OfficeColor labRed = OfficeColorSpaceConverter.FromLab(53.24D, 80.09D, 67.2D);
        OfficeColor cmykRed = OfficeColorSpaceConverter.FromCmyk(0D, 1D, 1D, 0D);

        Assert.InRange(labRed.R, 245, 255);
        Assert.InRange(labRed.G, 0, 15);
        Assert.InRange(labRed.B, 0, 15);
        Assert.Equal(OfficeColor.Red, cmykRed);
    }

    [Fact]
    public void IccMatrixProfile_ParsesAndConvertsThroughEmbeddedTrcs() {
        byte[] profileBytes = PdfIccProfiles.SrgbIec6196621;

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.NotNull(profile);
        Assert.Equal(3, profile!.ComponentCount);
        Assert.True(profile.TryConvert(new[] { 0.25D, 0.5D, 0.75D }, out OfficeColor color));
        Assert.InRange(color.R, 65, 75);
        Assert.InRange(color.G, 120, 140);
        Assert.InRange(color.B, 185, 205);
    }

    [Fact]
    public void IccMatrixProfile_RejectsUnsupportedOrMalformedProfiles() {
        byte[] cmykProfile = PdfIccProfiles.SrgbIec6196621;
        cmykProfile[16] = (byte)'C';
        cmykProfile[17] = (byte)'M';
        cmykProfile[18] = (byte)'Y';
        cmykProfile[19] = (byte)'K';
        byte[] badSignature = PdfIccProfiles.SrgbIec6196621;
        badSignature[36] = (byte)'x';
        byte[] authoredLutTransform = PdfIccProfiles.SrgbIec6196621;
        RenameTag(authoredLutTransform, "desc", "A2B0");
        byte[] outputProfile = PdfIccProfiles.SrgbIec6196621;
        WriteSignature(outputProfile, 12, "prtr");
        byte[] deviceLinkProfile = PdfIccProfiles.SrgbIec6196621;
        WriteSignature(deviceLinkProfile, 12, "link");

        Assert.False(OfficeIccColorProfile.TryCreate(cmykProfile, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(badSignature, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(authoredLutTransform, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(outputProfile, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(deviceLinkProfile, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(null!, out _));
    }

    [Fact]
    public void IccMatrixProfile_AppliesParametricCurves() {
        byte[] profileBytes = PdfIccProfiles.SrgbIec6196621;
        int curveOffset = FindTagOffset(profileBytes, "rTRC");
        WriteSignature(profileBytes, curveOffset, "para");
        Array.Clear(profileBytes, curveOffset + 4, 8);
        WriteS15Fixed16(profileBytes, curveOffset + 12, 2D);

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryConvert(new[] { 0.5D, 0.5D, 0.5D }, out OfficeColor color));
        Assert.InRange(color.R, 130, 145);
        Assert.InRange(color.G, 130, 145);
        Assert.InRange(color.B, 130, 145);
    }

    [Fact]
    public void IccGrayProfile_UsesGrayTrcAndMediaWhitePoint() {
        byte[] profileBytes = PdfIccProfiles.SrgbIec6196621;
        WriteSignature(profileBytes, 16, "GRAY");
        RenameTag(profileBytes, "rTRC", "kTRC");

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.Equal(1, profile!.ComponentCount);
        Assert.True(profile.TryConvert(new[] { 0.5D }, out OfficeColor color));
        Assert.InRange(Math.Abs(color.R - color.G), 0, 2);
        Assert.InRange(Math.Abs(color.G - color.B), 0, 2);
        Assert.InRange(color.R, 120, 140);
    }

    [Fact]
    public void IccGrayProfile_RequiresValidMediaWhitePoint() {
        byte[] missingWhitePoint = PdfIccProfiles.SrgbIec6196621;
        WriteSignature(missingWhitePoint, 16, "GRAY");
        RenameTag(missingWhitePoint, "rTRC", "kTRC");
        RenameTag(missingWhitePoint, "wtpt", "desc");
        byte[] nonpositiveWhitePoint = PdfIccProfiles.SrgbIec6196621;
        WriteSignature(nonpositiveWhitePoint, 16, "GRAY");
        RenameTag(nonpositiveWhitePoint, "rTRC", "kTRC");
        WriteS15Fixed16(nonpositiveWhitePoint, FindTagOffset(nonpositiveWhitePoint, "wtpt") + 8, 0D);

        Assert.False(OfficeIccColorProfile.TryCreate(missingWhitePoint, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(nonpositiveWhitePoint, out _));
    }

    private static int FindTagOffset(byte[] profile, string signature) {
        uint target = ToSignature(signature);
        int count = checked((int)ReadUInt32(profile, 128));
        for (int index = 0; index < count; index++) {
            int entry = 132 + index * 12;
            if (ReadUInt32(profile, entry) == target) return checked((int)ReadUInt32(profile, entry + 4));
        }
        throw new InvalidOperationException("ICC tag was not found: " + signature + ".");
    }

    private static void RenameTag(byte[] profile, string oldSignature, string newSignature) {
        uint target = ToSignature(oldSignature);
        int count = checked((int)ReadUInt32(profile, 128));
        for (int index = 0; index < count; index++) {
            int entry = 132 + index * 12;
            if (ReadUInt32(profile, entry) != target) continue;
            WriteSignature(profile, entry, newSignature);
            return;
        }
        throw new InvalidOperationException("ICC tag was not found: " + oldSignature + ".");
    }

    private static uint ToSignature(string value) =>
        ((uint)value[0] << 24) | ((uint)value[1] << 16) | ((uint)value[2] << 8) | value[3];

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        unchecked(((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3]);

    private static void WriteSignature(byte[] bytes, int offset, string signature) {
        for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)signature[index];
    }

    private static void WriteS15Fixed16(byte[] bytes, int offset, double value) {
        uint encoded = unchecked((uint)(int)Math.Round(value * 65536D));
        bytes[offset] = (byte)(encoded >> 24);
        bytes[offset + 1] = (byte)(encoded >> 16);
        bytes[offset + 2] = (byte)(encoded >> 8);
        bytes[offset + 3] = (byte)encoded;
    }
}
