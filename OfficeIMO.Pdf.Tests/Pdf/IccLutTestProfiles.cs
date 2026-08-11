namespace OfficeIMO.Tests.Pdf;

internal static class IccLutTestProfiles {
    private const double WhiteX = 0.9642D;
    private const double WhiteY = 1D;
    private const double WhiteZ = 0.8249D;
    private const double PcsXyzScale = 65535D / 32768D;

    internal static byte[] CreateCmykLut8() => Create("CMYK", 4, precision: 1, pcsIsLab: true);

    internal static byte[] CreateRgbLut16() => Create("RGB ", 3, precision: 2);

    internal static byte[] CreateRgbLabLut16() => Create("RGB ", 3, precision: 2, pcsIsLab: true);

    internal static byte[] CreateCmykXyzLut8() => Create("CMYK", 4, precision: 1);

    internal static byte[] CreateCmykLut8WithDistinctRelativeIntent() =>
        Create("CMYK", 4, precision: 1, pcsIsLab: true, includeDistinctRelativeIntent: true);

    private static byte[] Create(
        string colorSpace,
        int inputChannels,
        int precision,
        bool pcsIsLab = false,
        bool includeDistinctRelativeIntent = false) {
        const int gridPoints = 2;
        int inputEntries = precision == 1 ? 256 : 2;
        int outputEntries = precision == 1 ? 256 : 2;
        int tableOffset = precision == 1 ? 48 : 52;
        int gridSamples = 1 << inputChannels;
        int tagLength = tableOffset + inputChannels * inputEntries * precision + gridSamples * 3 * precision + 3 * outputEntries * precision;
        int paddedTagLength = (tagLength + 3) & ~3;
        const int tagOffset = 156;
        int secondTagOffset = tagOffset + paddedTagLength;
        var profile = new byte[tagOffset + paddedTagLength * (includeDistinctRelativeIntent ? 2 : 1)];
        WriteUInt32(profile, 0, (uint)profile.Length);
        WriteSignature(profile, 12, "scnr");
        WriteSignature(profile, 16, colorSpace);
        WriteSignature(profile, 20, pcsIsLab ? "Lab " : "XYZ ");
        WriteSignature(profile, 36, "acsp");
        WriteS15Fixed16(profile, 68, WhiteX);
        WriteS15Fixed16(profile, 72, WhiteY);
        WriteS15Fixed16(profile, 76, WhiteZ);
        WriteUInt32(profile, 128, includeDistinctRelativeIntent ? 2U : 1U);
        WriteSignature(profile, 132, "A2B0");
        WriteUInt32(profile, 136, tagOffset);
        WriteUInt32(profile, 140, (uint)tagLength);
        if (includeDistinctRelativeIntent) {
            WriteSignature(profile, 144, "A2B1");
            WriteUInt32(profile, 148, (uint)secondTagOffset);
            WriteUInt32(profile, 152, (uint)tagLength);
        }

        WriteSignature(profile, tagOffset, precision == 1 ? "mft1" : "mft2");
        profile[tagOffset + 8] = (byte)inputChannels;
        profile[tagOffset + 9] = 3;
        profile[tagOffset + 10] = gridPoints;
        for (int diagonal = 0; diagonal < 3; diagonal++) {
            WriteS15Fixed16(profile, tagOffset + 12 + (diagonal * 3 + diagonal) * 4, 1D);
        }
        if (precision == 2) {
            WriteUInt16(profile, tagOffset + 48, (ushort)inputEntries);
            WriteUInt16(profile, tagOffset + 50, (ushort)outputEntries);
        }

        int inputOffset = tagOffset + tableOffset;
        for (int channel = 0; channel < inputChannels; channel++) {
            WriteIdentityTable(profile, inputOffset + channel * inputEntries * precision, inputEntries, precision);
        }

        int clutOffset = inputOffset + inputChannels * inputEntries * precision;
        for (int index = 0; index < gridSamples; index++) {
            var components = new double[inputChannels];
            int coordinate = index;
            for (int channel = inputChannels - 1; channel >= 0; channel--) {
                components[channel] = coordinate & 1;
                coordinate >>= 1;
            }
            double red;
            double green;
            double blue;
            if (inputChannels == 4) {
                double blackFactor = 1D - components[3];
                red = (1D - components[0]) * blackFactor;
                green = (1D - components[1]) * blackFactor;
                blue = (1D - components[2]) * blackFactor;
            } else {
                red = components[0];
                green = components[1];
                blue = components[2];
            }
            double x = 0.4361D * red + 0.3851D * green + 0.1431D * blue;
            double y = 0.2225D * red + 0.7169D * green + 0.0606D * blue;
            double z = 0.0139D * red + 0.0971D * green + 0.7141D * blue;
            int valueOffset = clutOffset + index * 3 * precision;
            if (pcsIsLab) {
                XyzToLab(x, y, z, out double lightness, out double a, out double b);
                if (precision == 1) {
                    WriteNormalized(profile, valueOffset, lightness / 100D, precision);
                    WriteNormalized(profile, valueOffset + precision, (a + 128D) / 255D, precision);
                    WriteNormalized(profile, valueOffset + precision * 2, (b + 128D) / 255D, precision);
                } else {
                    WriteUInt16(profile, valueOffset, EncodeLegacyLab16(lightness / 100D, 65280D));
                    WriteUInt16(profile, valueOffset + precision, EncodeLegacyLab16(a + 128D, 256D));
                    WriteUInt16(profile, valueOffset + precision * 2, EncodeLegacyLab16(b + 128D, 256D));
                }
            } else {
                WriteNormalized(profile, valueOffset, x / PcsXyzScale, precision);
                WriteNormalized(profile, valueOffset + precision, y / PcsXyzScale, precision);
                WriteNormalized(profile, valueOffset + precision * 2, z / PcsXyzScale, precision);
            }
        }

        int outputOffset = clutOffset + gridSamples * 3 * precision;
        for (int channel = 0; channel < 3; channel++) {
            WriteIdentityTable(profile, outputOffset + channel * outputEntries * precision, outputEntries, precision);
        }
        if (includeDistinctRelativeIntent) {
            Buffer.BlockCopy(profile, tagOffset, profile, secondTagOffset, tagLength);
            profile[secondTagOffset + tableOffset + inputChannels * inputEntries * precision] ^= 1;
        }
        return profile;
    }

    private static void XyzToLab(double x, double y, double z, out double lightness, out double a, out double b) {
        double fx = LabFunction(x / WhiteX);
        double fy = LabFunction(y / WhiteY);
        double fz = LabFunction(z / WhiteZ);
        lightness = Math.Max(0D, 116D * fy - 16D);
        a = 500D * (fx - fy);
        b = 200D * (fy - fz);
    }

    private static double LabFunction(double value) => value > 216D / 24389D
        ? Math.Pow(value, 1D / 3D)
        : (841D / 108D * value) + (4D / 29D);

    private static void WriteIdentityTable(byte[] bytes, int offset, int entries, int precision) {
        for (int entry = 0; entry < entries; entry++) {
            WriteNormalized(bytes, offset + entry * precision, entry / (double)(entries - 1), precision);
        }
    }

    private static ushort EncodeLegacyLab16(double value, double scale) {
        double encoded = value * scale;
        return (ushort)Math.Round(encoded < 0D ? 0D : encoded > ushort.MaxValue ? ushort.MaxValue : encoded);
    }

    private static void WriteNormalized(byte[] bytes, int offset, double value, int precision) {
        double clamped = value < 0D ? 0D : value > 1D ? 1D : value;
        if (precision == 1) {
            bytes[offset] = (byte)Math.Round(clamped * 255D);
        } else {
            WriteUInt16(bytes, offset, (ushort)Math.Round(clamped * 65535D));
        }
    }

    private static void WriteSignature(byte[] bytes, int offset, string signature) {
        for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)signature[index];
    }

    private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)(value >> 8);
        bytes[offset + 1] = (byte)value;
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static void WriteS15Fixed16(byte[] bytes, int offset, double value) =>
        WriteUInt32(bytes, offset, unchecked((uint)(int)Math.Round(value * 65536D)));
}
