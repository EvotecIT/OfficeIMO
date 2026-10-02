using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeAv1StillFrameReader {
    private static void ReadQuantization(OfficeAv1Bits bits, OfficeAv1StillSequence sequence, OfficeAv1StillFrame frame) {
        frame.BaseQIndex = bits.Read(8);
        frame.DeltaQYDc = ReadDeltaQ(bits);
        if (!sequence.Monochrome) {
            bool differentUv = sequence.SeparateUvDeltaQ && bits.Flag();
            frame.DeltaQUDc = ReadDeltaQ(bits);
            frame.DeltaQUAc = ReadDeltaQ(bits);
            frame.DeltaQVDc = differentUv ? ReadDeltaQ(bits) : frame.DeltaQUDc;
            frame.DeltaQVAc = differentUv ? ReadDeltaQ(bits) : frame.DeltaQUAc;
        }
        frame.UsingQMatrix = bits.Flag();
        if (frame.UsingQMatrix) {
            frame.QMatrixLevels[0] = bits.Read(4);
            frame.QMatrixLevels[1] = bits.Read(4); // Still signaled for monochrome.
            frame.QMatrixLevels[2] = sequence.SeparateUvDeltaQ ? bits.Read(4) : frame.QMatrixLevels[1];
        }
    }

    private static int ReadDeltaQ(OfficeAv1Bits bits) => bits.Flag() ? bits.Signed(7) : 0;

    private static void ReadSegmentation(OfficeAv1Bits bits, OfficeAv1StillFrame frame) {
        int[] widths = { 8, 6, 6, 6, 6, 3, 0, 0 };
        int[] limits = { 255, 63, 63, 63, 63, 7, 0, 0 };
        frame.SegmentationEnabled = bits.Flag();
        // PRIMARY_REF_NONE implies update_map=1, temporal_update=0, update_data=1.
        if (frame.SegmentationEnabled) {
            for (int segment = 0; segment < 8; segment++) {
                for (int feature = 0; feature < 8; feature++) {
                    if (!bits.Flag()) continue;
                    frame.SegmentFeatures[segment, feature] = true;
                    int value = feature < 5 ? bits.Signed(widths[feature] + 1) : bits.Read(widths[feature]);
                    frame.SegmentData[segment, feature] = Math.Max(-limits[feature], Math.Min(limits[feature], value));
                    frame.LastActiveSegmentId = segment;
                    frame.SegmentIdPreSkip |= feature >= 5;
                }
            }
        }
        frame.CodedLossless = true;
        bool zeroDeltas = frame.DeltaQYDc == 0 && frame.DeltaQUDc == 0 && frame.DeltaQUAc == 0 &&
            frame.DeltaQVDc == 0 && frame.DeltaQVAc == 0;
        for (int segment = 0; segment < 8; segment++) {
            int q = Math.Max(0, Math.Min(255, frame.BaseQIndex + (frame.SegmentFeatures[segment, 0] ? frame.SegmentData[segment, 0] : 0)));
            frame.LosslessSegments[segment] = q == 0 && zeroDeltas;
            frame.CodedLossless &= frame.LosslessSegments[segment];
        }
        frame.AllLossless = frame.CodedLossless && frame.Width == frame.UpscaledWidth;
    }

    private static void ReadDeltaParameters(OfficeAv1Bits bits, OfficeAv1StillFrame frame) {
        frame.DeltaQPresent = frame.BaseQIndex > 0 && bits.Flag();
        if (frame.DeltaQPresent) frame.DeltaQResolution = bits.Read(2);
        frame.DeltaLoopFilterPresent = frame.DeltaQPresent && !frame.AllowIntraBlockCopy && bits.Flag();
        if (frame.DeltaLoopFilterPresent) {
            frame.DeltaLoopFilterResolution = bits.Read(2);
            frame.DeltaLoopFilterMulti = bits.Flag();
        }
    }

    private static void ReadFilters(OfficeAv1Bits bits, OfficeAv1StillSequence sequence, OfficeAv1StillFrame frame) {
        if (!frame.CodedLossless && !frame.AllowIntraBlockCopy) {
            frame.LoopFilterLevels[0] = bits.Read(6);
            frame.LoopFilterLevels[1] = bits.Read(6);
            if (!sequence.Monochrome && (frame.LoopFilterLevels[0] != 0 || frame.LoopFilterLevels[1] != 0)) {
                frame.LoopFilterLevels[2] = bits.Read(6);
                frame.LoopFilterLevels[3] = bits.Read(6);
            }
            frame.LoopFilterSharpness = bits.Read(3);
            frame.LoopFilterDeltaEnabled = bits.Flag();
            if (frame.LoopFilterDeltaEnabled && bits.Flag()) {
                for (int i = 0; i < 8; i++) if (bits.Flag()) frame.LoopFilterReferenceDeltas[i] = bits.Signed(7);
                for (int i = 0; i < 2; i++) if (bits.Flag()) frame.LoopFilterModeDeltas[i] = bits.Signed(7);
            }
        }
        if (!frame.CodedLossless && !frame.AllowIntraBlockCopy && sequence.Cdef) {
            frame.CdefDamping = bits.Read(2) + 3;
            frame.CdefBits = bits.Read(2);
            for (int i = 0; i < (1 << frame.CdefBits); i++) {
                frame.CdefStrengths[i, 0] = bits.Read(4);
                frame.CdefStrengths[i, 1] = ReadCdefSecondary(bits);
                if (!sequence.Monochrome) {
                    frame.CdefStrengths[i, 2] = bits.Read(4);
                    frame.CdefStrengths[i, 3] = ReadCdefSecondary(bits);
                }
            }
        }
        if (!frame.AllLossless && !frame.AllowIntraBlockCopy && sequence.Restoration) {
            bool usesRestoration = false, usesChroma = false;
            for (int i = 0; i < (sequence.Monochrome ? 1 : 3); i++) {
                frame.RestorationTypes[i] = bits.Read(2);
                usesRestoration |= frame.RestorationTypes[i] != 0;
                usesChroma |= i > 0 && frame.RestorationTypes[i] != 0;
            }
            if (usesRestoration) {
                int shift = bits.Read(1);
                if (sequence.Use128Superblock) shift++;
                else if (shift != 0) shift += bits.Read(1);
                int uvShift = usesChroma ? bits.Read(1) : 0; // This sequence path is always 4:2:0 or monochrome.
                frame.RestorationUnitSizes[0] = 256 >> (2 - shift);
                frame.RestorationUnitSizes[1] = frame.RestorationUnitSizes[0] >> uvShift;
                frame.RestorationUnitSizes[2] = frame.RestorationUnitSizes[1];
            }
        }
    }

    private static int ReadCdefSecondary(OfficeAv1Bits bits) {
        int strength = bits.Read(2);
        return strength == 3 ? 4 : strength;
    }
}
