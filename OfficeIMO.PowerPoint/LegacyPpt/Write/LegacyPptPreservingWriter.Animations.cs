using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordAnimationInfo = 0x1014;
        private const ushort RecordInteractiveInfoForAnimationOrder = 0x0FF2;
        private const ushort RecordPlaceholderForAnimationOrder = 0x0BC3;
        private const ushort RecordRecolorInfoForAnimationOrder = 0x0FE7;
        private const ushort RecordProgTagsForAnimationOrder = 0x1388;

        private static bool AnimationsEqual(LegacyPptAnimation? source,
            LegacyPptWriter.LegacyPptWriterAnimation? current) {
            if (source == null || current == null) return source == null && current == null;
            if ((byte)source.Effect != (byte)current.Effect
                || source.EffectDirection != current.Direction
                || (byte)source.BuildType != (byte)current.BuildType
                || source.Automatic != current.Automatic
                || source.Automatic
                && source.DelayMilliseconds != current.DelayMilliseconds
                || source.Order != current.Order
                || source.Reverse != current.Reverse
                || source.AnimateBackground != current.AnimateBackground
                || (byte)source.AfterEffect != (byte)current.AfterEffect
                || (byte)source.TextBuildSubEffect != (byte)current.TextBuild
                || source.RawDimColor != current.RawDimColor
                || source.StopsSound != current.StopsSound
                || source.PlaysSound != (current.Sound != null)) return false;
            return !source.PlaysSound
                || source.SoundIdReference == current.Sound!.Id;
        }

        private static bool TryRewriteClientDataAnimation(
            LegacyPptRecord clientData,
            LegacyPptWriter.LegacyPptWriterAnimation? append,
            out byte[] bytes) {
            if (clientData.Version != 0x0F
                || clientData.Children.Count(record =>
                    record.Type == RecordAnimationInfo) > 1) {
                bytes = clientData.CopyRecordBytes();
                return false;
            }
            var children = new List<byte[]>(clientData.Children.Count + 1);
            bool inserted = false;
            foreach (LegacyPptRecord child in clientData.Children) {
                if (child.Type == RecordAnimationInfo) continue;
                if (!inserted && append != null
                    && IsAfterAnimationInfo(child.Type)) {
                    children.Add(LegacyPptWriter.BuildAnimationInfoRecord(append));
                    inserted = true;
                }
                children.Add(child.CopyRecordBytes());
            }
            if (!inserted && append != null) {
                children.Add(LegacyPptWriter.BuildAnimationInfoRecord(append));
            }
            bytes = BuildRecord(clientData.Version, clientData.Instance,
                clientData.Type, Concat(children));
            return true;
        }

        private static bool IsAfterAnimationInfo(ushort recordType) =>
            recordType == RecordInteractiveInfoForAnimationOrder
            || recordType == RecordPlaceholderForAnimationOrder
            || recordType == RecordRecolorInfoForAnimationOrder
            || recordType == RecordProgTagsForAnimationOrder;
    }
}
