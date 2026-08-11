namespace OfficeIMO.Tests.Pdf;

internal static class IccMabTestProfiles {
    private const double WhiteX = 0.9642D;
    private const double WhiteY = 1D;
    private const double WhiteZ = 0.8249D;

    internal static byte[] CreateCmykLab8() =>
        Create("CMYK", 4, precision: 1, pcsIsLab: true, transformedStages: false, includeMatrix: false, includeClut: true);

    internal static byte[] CreateCmykLab8WithSharedCurves() =>
        Create(
            "CMYK",
            4,
            precision: 1,
            pcsIsLab: true,
            transformedStages: false,
            includeMatrix: false,
            includeClut: true,
            shareInputAndOutputCurves: true);

    internal static byte[] CreateRgbXyz16WithTransformedStages() =>
        Create("RGB ", 3, precision: 2, pcsIsLab: false, transformedStages: true, includeMatrix: true, includeClut: true);

    internal static byte[] CreateRgbXyzBOnly() =>
        Create("RGB ", 3, precision: 2, pcsIsLab: false, transformedStages: false, includeMatrix: false, includeClut: false);

    internal static byte[] CreateRgbXyzMatrixOnly() =>
        Create("RGB ", 3, precision: 2, pcsIsLab: false, transformedStages: true, includeMatrix: true, includeClut: false);

    internal static int FindTransformOffset(byte[] profile) => checked((int)ReadUInt32(profile, 136));

    private static byte[] Create(
        string colorSpace,
        int inputChannels,
        int precision,
        bool pcsIsLab,
        bool transformedStages,
        bool includeMatrix,
        bool includeClut,
        bool shareInputAndOutputCurves = false) {
        const int tagOffset = 156;
        const int bOffset = 32;
        int storedBCurveCount = shareInputAndOutputCurves ? Math.Max(3, inputChannels) : 3;
        int cursor = bOffset + storedBCurveCount * 16;
        int matrixOffset = includeMatrix ? cursor : 0;
        if (includeMatrix) cursor += 48;
        int mOffset = includeMatrix ? cursor : 0;
        if (includeMatrix) cursor += 3 * 16;
        int clutOffset = includeClut ? cursor : 0;
        if (includeClut) {
            int gridSamples = GetGridSampleCount(inputChannels, variableGrid: inputChannels == 4);
            cursor += Align4(20 + gridSamples * 3 * precision);
        }
        int aOffset = includeClut ? (shareInputAndOutputCurves ? bOffset : cursor) : 0;
        if (includeClut && !shareInputAndOutputCurves) cursor += inputChannels * 16;
        int tagLength = cursor;
        var profile = new byte[tagOffset + tagLength];

        WriteUInt32(profile, 0, (uint)profile.Length);
        profile[8] = 0x04;
        profile[9] = 0x40;
        WriteSignature(profile, 12, "scnr");
        WriteSignature(profile, 16, colorSpace);
        WriteSignature(profile, 20, pcsIsLab ? "Lab " : "XYZ ");
        WriteSignature(profile, 36, "acsp");
        WriteS15Fixed16(profile, 68, WhiteX);
        WriteS15Fixed16(profile, 72, WhiteY);
        WriteS15Fixed16(profile, 76, WhiteZ);
        WriteUInt32(profile, 128, 1);
        WriteSignature(profile, 132, "A2B0");
        WriteUInt32(profile, 136, tagOffset);
        WriteUInt32(profile, 140, (uint)tagLength);

        WriteSignature(profile, tagOffset, "mAB ");
        profile[tagOffset + 8] = (byte)inputChannels;
        profile[tagOffset + 9] = 3;
        WriteUInt32(profile, tagOffset + 12, bOffset);
        WriteUInt32(profile, tagOffset + 16, (uint)matrixOffset);
        WriteUInt32(profile, tagOffset + 20, (uint)mOffset);
        WriteUInt32(profile, tagOffset + 24, (uint)clutOffset);
        WriteUInt32(profile, tagOffset + 28, (uint)aOffset);

        WriteCurveSet(
            profile,
            tagOffset + bOffset,
            storedBCurveCount,
            transformedStages ? 1.25D : 1D,
            transformedStages ? 0.25D : 0D,
            sampled: !includeMatrix && !includeClut);
        if (includeMatrix) {
            WriteMatrix(profile, tagOffset + matrixOffset, transformedStages);
            WriteCurveSet(
                profile,
                tagOffset + mOffset,
                3,
                transformedStages ? 1.5D : 1D,
                transformedStages ? 0.25D : 0D,
                sampled: false);
        }
        if (includeClut) {
            WriteClut(profile, tagOffset + clutOffset, inputChannels, precision, pcsIsLab, transformedStages);
            if (!shareInputAndOutputCurves) {
                WriteCurveSet(
                    profile,
                    tagOffset + aOffset,
                    inputChannels,
                    transformedStages ? 2D : 1D,
                    transformedStages ? 0.5D : 0D,
                    sampled: false);
            }
        }
        return profile;
    }

    private static void WriteCurveSet(
        byte[] bytes,
        int offset,
        int count,
        double gamma,
        double gammaStep,
        bool sampled) {
        for (int index = 0; index < count; index++) {
            int curveOffset = offset + index * 16;
            if (sampled) {
                WriteSignature(bytes, curveOffset, "curv");
                WriteUInt32(bytes, curveOffset + 8, 2);
                WriteUInt16(bytes, curveOffset + 12, 0);
                WriteUInt16(bytes, curveOffset + 14, ushort.MaxValue);
            } else {
                WriteSignature(bytes, curveOffset, "para");
                WriteS15Fixed16(bytes, curveOffset + 12, gamma + index * gammaStep);
            }
        }
    }

    private static void WriteMatrix(byte[] bytes, int offset, bool transformedStages) {
        for (int diagonal = 0; diagonal < 3; diagonal++) {
            WriteS15Fixed16(bytes, offset + (diagonal * 3 + diagonal) * 4, transformedStages ? 0.5D : 1D);
        }
        if (transformedStages) {
            WriteS15Fixed16(bytes, offset + 36, 0.1D);
            WriteS15Fixed16(bytes, offset + 40, 0.1D);
            WriteS15Fixed16(bytes, offset + 44, 0.1D);
        }
    }

    private static void WriteClut(
        byte[] bytes,
        int offset,
        int inputChannels,
        int precision,
        bool pcsIsLab,
        bool transformedStages) {
        bool variableGrid = inputChannels == 4;
        for (int channel = 0; channel < inputChannels; channel++) bytes[offset + channel] = (byte)GetGridPoints(channel, variableGrid);
        bytes[offset + 16] = (byte)precision;
        int gridSamples = GetGridSampleCount(inputChannels, variableGrid);
        for (int index = 0; index < gridSamples; index++) {
            var components = new double[inputChannels];
            int coordinate = index;
            for (int channel = inputChannels - 1; channel >= 0; channel--) {
                int gridPoints = GetGridPoints(channel, variableGrid);
                components[channel] = coordinate % gridPoints / (double)(gridPoints - 1);
                coordinate /= gridPoints;
            }

            double output0;
            double output1;
            double output2;
            if (transformedStages) {
                output0 = components[0] * (0.25D + 0.75D * components[1]);
                output1 = components[1] * (0.25D + 0.75D * components[2]);
                output2 = components[2] * (0.25D + 0.75D * components[0]);
            } else {
                double blackFactor = inputChannels == 4 ? 1D - components[3] : 1D;
                double red = (1D - components[0]) * blackFactor;
                double green = (1D - components[1]) * blackFactor;
                double blue = (1D - components[2]) * blackFactor;
                double x = 0.4361D * red + 0.3851D * green + 0.1431D * blue;
                double y = 0.2225D * red + 0.7169D * green + 0.0606D * blue;
                double z = 0.0139D * red + 0.0971D * green + 0.7141D * blue;
                if (pcsIsLab) {
                    XyzToLab(x, y, z, out double lightness, out double a, out double b);
                    output0 = lightness / 100D;
                    output1 = (a + 128D) / 255D;
                    output2 = (b + 128D) / 255D;
                } else {
                    const double pcsXyzScale = 65535D / 32768D;
                    output0 = x / pcsXyzScale;
                    output1 = y / pcsXyzScale;
                    output2 = z / pcsXyzScale;
                }
            }

            int valueOffset = offset + 20 + index * 3 * precision;
            WriteNormalized(bytes, valueOffset, output0, precision);
            WriteNormalized(bytes, valueOffset + precision, output1, precision);
            WriteNormalized(bytes, valueOffset + precision * 2, output2, precision);
        }
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

    private static void WriteNormalized(byte[] bytes, int offset, double value, int precision) {
        double clamped = value < 0D ? 0D : value > 1D ? 1D : value;
        if (precision == 1) {
            bytes[offset] = (byte)Math.Round(clamped * 255D);
        } else {
            WriteUInt16(bytes, offset, (ushort)Math.Round(clamped * 65535D));
        }
    }

    private static int GetGridSampleCount(int inputChannels, bool variableGrid) {
        int count = 1;
        for (int channel = 0; channel < inputChannels; channel++) count *= GetGridPoints(channel, variableGrid);
        return count;
    }

    private static int GetGridPoints(int channel, bool variableGrid) => variableGrid && channel == 1 ? 3 : 2;

    private static int Align4(int value) => (value + 3) & ~3;

    private static void WriteSignature(byte[] bytes, int offset, string signature) {
        for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)signature[index];
    }

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];

    internal static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)(value >> 8);
        bytes[offset + 1] = (byte)value;
    }

    internal static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static void WriteS15Fixed16(byte[] bytes, int offset, double value) =>
        WriteUInt32(bytes, offset, unchecked((uint)(int)Math.Round(value * 65536D)));
}
