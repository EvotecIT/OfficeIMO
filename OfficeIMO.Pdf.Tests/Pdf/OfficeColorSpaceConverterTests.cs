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

        Assert.False(OfficeIccColorProfile.TryCreate(cmykProfile, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(badSignature, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(authoredLutTransform, out _));
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
    public void IccLut8Profile_ConvertsCmykThroughMultidimensionalClut() {
        byte[] profileBytes = IccLutTestProfiles.CreateCmykLut8();

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.Equal(4, profile!.ComponentCount);
        Assert.True(profile.TryConvert(new[] { 0D, 1D, 1D, 0D }, out OfficeColor color));
        Assert.InRange(color.R, 245, 255);
        Assert.InRange(color.G, 0, 15);
        Assert.InRange(color.B, 0, 15);
    }

    [Fact]
    public void IccLut8Profile_InterpolatesInteriorCmykClutCoordinates() {
        byte[] profileBytes = IccLutTestProfiles.CreateCmykLut8();

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryConvert(new[] { 0.25D, 0.5D, 0.75D, 0.2D }, out OfficeColor color));
        Assert.InRange(color.R, 140, 175);
        Assert.InRange(color.G, 80, 125);
        Assert.InRange(color.B, 25, 80);
    }

    [Fact]
    public void IccLut16Profile_InterpolatesRgbInputAndOutputTables() {
        byte[] profileBytes = IccLutTestProfiles.CreateRgbLut16();

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.Equal(3, profile!.ComponentCount);
        Assert.True(profile.TryConvert(new[] { 0.5D, 0.25D, 0.75D }, out OfficeColor color));
        Assert.InRange(color.R, 180, 195);
        Assert.InRange(color.G, 125, 145);
        Assert.InRange(color.B, 220, 235);
    }

    [Fact]
    public void IccLut16Profile_ConvertsLabProfileConnectionSpace() {
        byte[] profileBytes = IccLutTestProfiles.CreateRgbLabLut16();

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryConvert(new[] { 1D, 0D, 0D }, out OfficeColor color));
        OfficeColor expected = OfficeColorSpaceConverter.FromXyz(
            0.4361D,
            0.2225D,
            0.0139D,
            0.9642D,
            1D,
            0.8249D);
        Assert.InRange(Math.Abs(color.R - expected.R), 0, 1);
        Assert.InRange(Math.Abs(color.G - expected.G), 0, 1);
        Assert.InRange(Math.Abs(color.B - expected.B), 0, 1);
    }

    [Fact]
    public void IccLut8Profile_RejectsUndefinedXyzProfileConnectionEncoding() {
        byte[] profileBytes = IccLutTestProfiles.CreateCmykXyzLut8();

        Assert.False(OfficeIccColorProfile.TryCreate(profileBytes, out _));
    }

    [Fact]
    public void IccLutProfile_RejectsNonIdentityDeviceMatrix() {
        byte[] profileBytes = IccLutTestProfiles.CreateCmykLut8();
        int transformOffset = FindTagOffset(profileBytes, "A2B0");
        WriteS15Fixed16(profileBytes, transformOffset + 16, 0.5D);

        Assert.False(OfficeIccColorProfile.TryCreate(profileBytes, out _));
    }

    [Fact]
    public void IccLutProfile_RejectsNonD50IlluminantAndUnexpectedTagPayload() {
        byte[] nonD50Profile = IccLutTestProfiles.CreateCmykLut8();
        WriteS15Fixed16(nonD50Profile, 68, 0.95D);
        byte[] oversizedTagProfile = IccLutTestProfiles.CreateCmykLut8();
        int tagEntry = FindTagEntryOffset(oversizedTagProfile, "A2B0");
        WriteUInt32(oversizedTagProfile, tagEntry + 8, ReadUInt32(oversizedTagProfile, tagEntry + 8) + 1U);

        Assert.False(OfficeIccColorProfile.TryCreate(nonD50Profile, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(oversizedTagProfile, out _));
    }

    [Fact]
    public void IccLutProfile_RejectsIntentDependentTransformsUntilIntentSelectionIsSupported() {
        byte[] profileBytes = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();

        Assert.False(OfficeIccColorProfile.TryCreate(profileBytes, out _));
    }

    [Fact]
    public void IccLutProfile_RejectsUnsupportedDToBTransformWithoutMatrixFallback() {
        byte[] profileBytes = IccLutTestProfiles.CreateRgbLut16();
        WriteSignature(profileBytes, FindTagEntryOffset(profileBytes, "A2B0"), "D2B0");

        Assert.False(OfficeIccColorProfile.TryCreate(profileBytes, out _));
    }

    [Fact]
    public void IccMabProfile_ConvertsCmykThroughVariableGridClut() {
        byte[] profileBytes = IccMabTestProfiles.CreateCmykLab8();

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.Equal(4, profile!.ComponentCount);
        Assert.True(profile.TryConvert(new[] { 0D, 1D, 1D, 0D }, out OfficeColor color));
        Assert.InRange(color.R, 245, 255);
        Assert.InRange(color.G, 0, 15);
        Assert.InRange(color.B, 0, 15);

        Assert.True(profile.TryConvert(new[] { 0.25D, 0.5D, 0.75D, 0.2D }, out OfficeColor interior));
        Assert.InRange(interior.R, 140, 175);
        Assert.InRange(interior.G, 130, 170);
        Assert.InRange(interior.B, 60, 100);
    }

    [Fact]
    public void IccMabProfile_AppliesCurvesClutMatrixAndOffsetsInSpecificationOrder() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithTransformedStages();

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryConvert(new[] { 0.5D, 0.25D, 0.75D }, out OfficeColor color));
        ApplyMabStages(0.5D, 0.25D, 0.75D, out double expectedX, out double expectedY, out double expectedZ);
        OfficeColor expected = OfficeColorSpaceConverter.FromXyz(
            expectedX * (65535D / 32768D),
            expectedY * (65535D / 32768D),
            expectedZ * (65535D / 32768D),
            0.9642D,
            1D,
            0.8249D);
        Assert.InRange(Math.Abs(color.R - expected.R), 0, 1);
        Assert.InRange(Math.Abs(color.G - expected.G), 0, 1);
        Assert.InRange(Math.Abs(color.B - expected.B), 0, 1);
    }

    [Fact]
    public void IccMabProfile_AcceptsAllSpecificationDefinedElementCombinations() {
        Assert.True(OfficeIccColorProfile.TryCreate(IccMabTestProfiles.CreateRgbXyzBOnly(), out _));
        Assert.True(OfficeIccColorProfile.TryCreate(IccMabTestProfiles.CreateRgbXyzMatrixOnly(), out _));
        Assert.True(OfficeIccColorProfile.TryCreate(IccMabTestProfiles.CreateCmykLab8(), out _));
        Assert.True(OfficeIccColorProfile.TryCreate(IccMabTestProfiles.CreateRgbXyz16WithTransformedStages(), out _));
    }

    [Fact]
    public void IccMabProfile_AllowsCurveSharingButRejectsMatrixCurveAliasing() {
        byte[] sharedCurves = IccMabTestProfiles.CreateCmykLab8WithSharedCurves();
        byte[] matrixCurveAlias = IccMabTestProfiles.CreateRgbXyzMatrixOnly();
        int tagOffset = IccMabTestProfiles.FindTransformOffset(matrixCurveAlias);
        int originalMatrixOffset = tagOffset + checked((int)ReadUInt32(matrixCurveAlias, tagOffset + 16));
        Array.Clear(matrixCurveAlias, originalMatrixOffset, 48);
        IccMabTestProfiles.WriteUInt32(
            matrixCurveAlias,
            tagOffset + 16,
            ReadUInt32(matrixCurveAlias, tagOffset + 20));

        Assert.True(OfficeIccColorProfile.TryCreate(sharedCurves, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryConvert(new[] { 0D, 1D, 1D, 0D }, out OfficeColor color));
        Assert.InRange(color.R, 245, 255);
        Assert.False(OfficeIccColorProfile.TryCreate(matrixCurveAlias, out _));
    }

    [Fact]
    public void IccMabProfile_RejectsMalformedElementLayoutsAndClutMetadata() {
        byte[] reserved = IccMabTestProfiles.CreateCmykLab8();
        byte[] misalignedMatrix = IccMabTestProfiles.CreateCmykLab8();
        byte[] missingClut = IccMabTestProfiles.CreateCmykLab8();
        byte[] unusedGridDimension = IccMabTestProfiles.CreateCmykLab8();
        byte[] unsupportedPrecision = IccMabTestProfiles.CreateCmykLab8();
        int tagOffset = IccMabTestProfiles.FindTransformOffset(reserved);
        reserved[tagOffset + 4] = 1;
        IccMabTestProfiles.WriteUInt32(misalignedMatrix, tagOffset + 16, 81);
        IccMabTestProfiles.WriteUInt32(missingClut, tagOffset + 24, 0);
        int clutOffset = tagOffset + checked((int)ReadUInt32(unusedGridDimension, tagOffset + 24));
        unusedGridDimension[clutOffset + 5] = 2;
        unsupportedPrecision[clutOffset + 16] = 3;

        Assert.False(OfficeIccColorProfile.TryCreate(reserved, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(misalignedMatrix, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(missingClut, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(unusedGridDimension, out _));
        Assert.False(OfficeIccColorProfile.TryCreate(unsupportedPrecision, out _));
    }

    [Fact]
    public void IccMabProfile_RejectsTransformDeclaredByPreV4Profile() {
        byte[] profileBytes = IccMabTestProfiles.CreateCmykLab8();
        profileBytes[8] = 3;
        profileBytes[9] = 0x40;

        Assert.False(OfficeIccColorProfile.TryCreate(profileBytes, out _));
    }

    [Fact]
    public void IccMabProfile_AcceptsBoundedZeroPaddingAndRejectsAuthoredTailData() {
        byte[] original = IccMabTestProfiles.CreateCmykLab8();
        const int tailLength = 1024 * 1024;
        var profileBytes = new byte[original.Length + tailLength];
        Buffer.BlockCopy(original, 0, profileBytes, 0, original.Length);
        IccMabTestProfiles.WriteUInt32(profileBytes, 0, (uint)profileBytes.Length);
        IccMabTestProfiles.WriteUInt32(
            profileBytes,
            140,
            checked(ReadUInt32(original, 140) + (uint)tailLength));

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out _));

        profileBytes[^1] = 1;
        Assert.False(OfficeIccColorProfile.TryCreate(profileBytes, out _));
    }

    private static void ApplyMabStages(
        double input0,
        double input1,
        double input2,
        out double output0,
        out double output1,
        out double output2) {
        double a0 = Math.Pow(input0, 2D);
        double a1 = Math.Pow(input1, 2.5D);
        double a2 = Math.Pow(input2, 3D);
        double clut0 = a0 * (0.25D + 0.75D * a1);
        double clut1 = a1 * (0.25D + 0.75D * a2);
        double clut2 = a2 * (0.25D + 0.75D * a0);
        double m0 = Math.Pow(clut0, 1.5D);
        double m1 = Math.Pow(clut1, 1.75D);
        double m2 = Math.Pow(clut2, 2D);
        output0 = Math.Pow(Math.Min(1D, 0.5D * m0 + 0.1D), 1.25D);
        output1 = Math.Pow(Math.Min(1D, 0.5D * m1 + 0.1D), 1.5D);
        output2 = Math.Pow(Math.Min(1D, 0.5D * m2 + 0.1D), 1.75D);
    }

    private static int FindTagOffset(byte[] profile, string signature) {
        int entry = FindTagEntryOffset(profile, signature);
        return checked((int)ReadUInt32(profile, entry + 4));
    }

    private static int FindTagEntryOffset(byte[] profile, string signature) {
        uint target = ToSignature(signature);
        int count = checked((int)ReadUInt32(profile, 128));
        for (int index = 0; index < count; index++) {
            int entry = 132 + index * 12;
            if (ReadUInt32(profile, entry) == target) return entry;
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

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }
}
