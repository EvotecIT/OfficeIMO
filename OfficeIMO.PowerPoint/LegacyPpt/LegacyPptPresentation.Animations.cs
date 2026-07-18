using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordAnimationInfoAtom = 0x0FF1;
        private const ushort RecordAnimationInfo = 0x1014;
        private const ushort OfficeArtClientDataForAnimation = 0xF011;

        private LegacyPptAnimation? ReadShapeAnimation(LegacyPptRecord shapeContainer,
            LegacyPptImportOptions options) {
            LegacyPptRecord[] containers = shapeContainer.Children
                .Where(record => record.Type == OfficeArtClientDataForAnimation)
                .SelectMany(record => record.Children)
                .Where(record => record.Type == RecordAnimationInfo)
                .ToArray();
            if (containers.Length == 0) return null;
            if (containers.Length != 1) {
                AddDiagnostic("PPT-ANIMATION-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A shape contains multiple classic animation containers; they remain preserve-only.",
                    shapeContainer.Offset);
                return null;
            }

            LegacyPptRecord container = containers[0];
            int soundOverrideCount = container.Children.Count(record =>
                record.Type == RecordSound);
            if (soundOverrideCount > 0) {
                AddDiagnostic("PPT-ANIMATION-SOUND-OVERRIDE",
                    LegacyPptDiagnosticSeverity.Warning,
                    soundOverrideCount == 1
                        ? "A classic animation contains an inline sound override; its atom and sound remain preservation-only."
                        : "A classic animation contains multiple inline sound overrides and remains preservation-only.",
                    container.Offset);
            }
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordAnimationInfoAtom).ToArray();
            if (container.Version != 0x0F || container.Instance != 0
                || atoms.Length != 1 || atoms[0].Version != 1
                || atoms[0].Instance != 0 || atoms[0].PayloadLength != 28) {
                AddDiagnostic("PPT-ANIMATION-ATOM",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A classic animation container or atom has an invalid record header and remains preserve-only.",
                    container.Offset);
                return null;
            }

            LegacyPptRecord atom = atoms[0];
            uint flags = atom.ReadUInt32(4);
            if (!HasValidTwoBitFlags(flags)) {
                AddDiagnostic("PPT-ANIMATION-FLAGS",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A classic animation uses an undefined two-bit flag value and remains preserve-only.",
                    atom.Offset);
                return null;
            }
            if (options.ReportUnsupportedContent && (flags & 0xFFFF0000U) != 0) {
                AddDiagnostic("PPT-ANIMATION-RESERVED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A classic animation has nonzero reserved flags; the raw bits remain preserved.",
                    atom.Offset);
            }

            int delay = atom.ReadInt32(12);
            short order = atom.ReadInt16(16);
            byte buildValue = atom.ReadByte(20);
            byte effectValue = atom.ReadByte(21);
            byte direction = atom.ReadByte(22);
            byte afterValue = atom.ReadByte(23);
            byte textBuildValue = atom.ReadByte(24);
            ushort rawUnused = atom.ReadUInt16(26);
            if (options.ReportUnsupportedContent && rawUnused != 0) {
                AddDiagnostic("PPT-ANIMATION-UNUSED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A classic animation has nonzero reserved trailing bytes; they remain preserved.",
                    atom.Offset);
            }
            if (order < -2 || ((flags >> 2) & 0x03U) == 1U && delay < 0
                || !IsAnimationBuildType(buildValue)
                || !IsAnimationEffect(effectValue)
                || !IsAnimationDirection(effectValue, direction)
                || afterValue > (byte)LegacyPptAnimationAfterEffect.HideImmediately
                || textBuildValue > (byte)LegacyPptTextBuildSubEffect.ByCharacter) {
                AddDiagnostic("PPT-ANIMATION-VALUE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A classic animation contains an invalid trigger, order, build, effect, direction, or after-effect value and remains preserve-only.",
                    atom.Offset);
                return null;
            }

            return new LegacyPptAnimation(atom.ReadUInt32(0), flags,
                atom.ReadUInt32(8), delay, order, atom.ReadUInt16(18),
                (LegacyPptAnimationBuildType)buildValue,
                (LegacyPptAnimationEffect)effectValue, direction,
                (LegacyPptAnimationAfterEffect)afterValue,
                (LegacyPptTextBuildSubEffect)textBuildValue,
                atom.ReadByte(25), rawUnused,
                hasSoundOverride: soundOverrideCount > 0);
        }

        private static bool HasValidTwoBitFlags(uint flags) {
            for (int shift = 0; shift < 16; shift += 2) {
                if (((flags >> shift) & 0x03U) > 1U) return false;
            }
            return true;
        }

        private static bool IsAnimationBuildType(byte value) => value <= 0x0A
            || value == (byte)LegacyPptAnimationBuildType.FollowMaster;

        private static bool IsAnimationEffect(byte value) => value <= 0x0E
            || value is 0x11 or 0x12 or 0x13 or 0x1A or 0x1B;

        private static bool IsAnimationDirection(byte effect, byte direction) => effect switch {
            0x00 => direction <= 2,
            0x01 => true,
            0x02 or 0x03 or 0x08 or 0x0B => direction <= 1,
            0x04 or 0x07 => direction <= 7,
            0x05 or 0x06 or 0x11 or 0x12 or 0x13 or 0x1B => direction == 0,
            0x09 => direction is >= 4 and <= 7,
            0x0A or 0x0D => direction <= 3,
            0x0C => direction <= 0x1C,
            0x0E => direction <= 2,
            0x1A => direction is 1 or 2 or 3 or 4 or 8,
            _ => false
        };
    }
}
