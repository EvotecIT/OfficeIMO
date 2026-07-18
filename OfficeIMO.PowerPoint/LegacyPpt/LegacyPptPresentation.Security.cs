using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        internal bool HasVbaContent { get; private set; }
        internal bool HasEmbeddedOleContent { get; private set; }
        internal bool HasLinkedOleContent { get; private set; }
        internal bool HasActiveXContent { get; private set; }
        internal bool HasExternalHyperlinkContent { get; private set; }
        internal bool HasExternalMediaContent { get; private set; }
        internal bool HasRunProgramContent { get; private set; }
        private LegacyPptRecordTraversalBudget _securityRecordBudget = null!;
        private readonly HashSet<int> _securityRecordOffsets = new();

        private void CaptureRawContentSecurityEvidence(
            LegacyPptRecord document) {
            foreach (RawRecordHeader owner in ReadRawChildHeaders(document)) {
                if (owner.Type == RecordExternalObjectList) {
                    CaptureRawExternalObjectEvidence(document, owner);
                } else if (owner.Type == RecordDocInfoList) {
                    CaptureRawVbaEvidence(document, owner);
                }
            }
        }

        private void CaptureRawExternalObjectEvidence(
            LegacyPptRecord document, RawRecordHeader list) {
            foreach (RawRecordHeader child in ReadRawChildHeaders(document,
                         list)) {
                switch (child.Type) {
                    case RecordExternalHyperlink:
                        if (HasReadableExternalHyperlinkTarget(document,
                                child)) {
                            HasExternalHyperlinkContent = true;
                        }
                        break;
                    case RecordExternalOleEmbed:
                        HasEmbeddedOleContent = true;
                        break;
                    case RecordExternalOleLink:
                        HasLinkedOleContent = true;
                        break;
                    case RecordExternalOleControl:
                        HasActiveXContent = true;
                        break;
                    case RecordExternalAviMovie:
                    case RecordExternalMciMovie:
                    case RecordExternalMidiAudio:
                    case RecordExternalWavAudioLink:
                        HasExternalMediaContent = true;
                        break;
                }
            }
        }

        private void CaptureRawVbaEvidence(LegacyPptRecord document,
            RawRecordHeader docInfo) {
            foreach (RawRecordHeader vbaInfo in ReadRawChildHeaders(document,
                         docInfo).Where(record => record.Type ==
                             RecordVbaInfo)) {
                if (HasMacroEvidence(document, vbaInfo)) {
                    HasVbaContent = true;
                }
            }
        }

        private bool HasMacroEvidence(LegacyPptRecord document,
            RawRecordHeader vbaInfo) {
            if (vbaInfo.Version != 0x0F || vbaInfo.Instance != 1) return true;
            RawRecordHeader atom = default;
            int atomCount = 0;
            foreach (RawRecordHeader child in ReadRawChildHeaders(document,
                         vbaInfo)) {
                if (child.Type != RecordVbaInfoAtom) continue;
                atom = child;
                atomCount++;
            }
            if (atomCount != 1 || atom.Version != 2 || atom.Instance != 0
                || atom.PayloadLength != 12) return true;
            return document.ReadUInt32(atom.PayloadOffset + 4) != 0;
        }

        private IEnumerable<RawRecordHeader> ReadRawChildHeaders(
            LegacyPptRecord container) => ReadRawChildHeaders(container,
            payloadOffset: 0, container.PayloadLength);

        private IEnumerable<RawRecordHeader> ReadRawChildHeaders(
            LegacyPptRecord container, RawRecordHeader parent) =>
            ReadRawChildHeaders(container, parent.PayloadOffset,
                parent.PayloadLength);

        private IEnumerable<RawRecordHeader> ReadRawChildHeaders(
            LegacyPptRecord container, int payloadOffset,
            int payloadLength) {
            if (payloadOffset < 0 || payloadLength < 0
                || payloadOffset > container.PayloadLength - payloadLength) {
                yield break;
            }
            int endOffset = payloadOffset + payloadLength;
            int position = payloadOffset;
            while (position <= endOffset - 8) {
                ushort versionAndInstance = container.ReadUInt16(position);
                uint declaredLength = container.ReadUInt32(position + 4);
                if (declaredLength > int.MaxValue) yield break;
                int childLength = unchecked((int)declaredLength);
                int childPayloadOffset = position + 8;
                if (childLength > endOffset - childPayloadOffset) yield break;
                int recordOffset = checked(container.PayloadOffset
                    + position);
                if (_securityRecordOffsets.Add(recordOffset)) {
                    _securityRecordBudget.Consume();
                }
                yield return new RawRecordHeader(
                    unchecked((byte)(versionAndInstance & 0x000F)),
                    unchecked((ushort)(versionAndInstance >> 4)),
                    container.ReadUInt16(position + 2),
                    childPayloadOffset, childLength);
                position = childPayloadOffset + childLength;
            }
        }

        private readonly struct RawRecordHeader {
            internal RawRecordHeader(byte version, ushort instance,
                ushort type, int payloadOffset, int payloadLength) {
                Version = version;
                Instance = instance;
                Type = type;
                PayloadOffset = payloadOffset;
                PayloadLength = payloadLength;
            }

            internal byte Version { get; }
            internal ushort Instance { get; }
            internal ushort Type { get; }
            internal int PayloadOffset { get; }
            internal int PayloadLength { get; }
        }
    }
}
