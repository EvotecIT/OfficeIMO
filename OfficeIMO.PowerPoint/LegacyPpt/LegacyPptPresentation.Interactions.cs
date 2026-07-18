using System.Collections.ObjectModel;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordExternalObjectList = 0x0409;
        private const ushort RecordExternalObjectListAtom = 0x040A;
        private const ushort RecordExternalHyperlinkAtom = 0x0FD3;
        private const ushort RecordExternalHyperlink = 0x0FD7;
        private const ushort RecordTextInteractiveInfoAtom = 0x0FDF;
        private const ushort RecordInteractiveInfo = 0x0FF2;
        private const ushort RecordInteractiveInfoAtom = 0x0FF3;
        private const ushort OfficeArtClientData = 0xF011;

        private readonly List<LegacyPptHyperlink> _hyperlinks = new();
        private readonly Dictionary<uint, LegacyPptHyperlink> _hyperlinksById = new();

        /// <summary>Gets document-level hyperlink targets by binary identifier.</summary>
        public IReadOnlyList<LegacyPptHyperlink> Hyperlinks => _hyperlinks;

        private void ParseHyperlinks(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            LegacyPptRecord[] lists = document.Children.Where(record =>
                record.Type == RecordExternalObjectList).ToArray();
            if (lists.Length == 0) {
                ParseHyperlinkExtensions(document, options);
                return;
            }
            if (lists.Length > 1) {
                AddDiagnostic("PPT-HYPERLINK-LISTS", LegacyPptDiagnosticSeverity.Warning,
                    "The document has multiple external-object lists; hyperlink edits remain loss-blocked.",
                    lists[1].Offset);
            }
            foreach (LegacyPptRecord list in lists) {
                ParseHyperlinkList(list, options);
            }
            ParseHyperlinkExtensions(document, options);
        }

        private void ParseHyperlinkList(LegacyPptRecord list,
            LegacyPptImportOptions options) {
            foreach (LegacyPptRecord container in list.Children.Where(
                         record => record.Type == RecordExternalHyperlink)) {
                if (HasReadableExternalHyperlinkTarget(container)) {
                    HasExternalHyperlinkContent = true;
                }
            }
            if (list.Version != 0x0F || list.Instance != 0) {
                AddDiagnostic("PPT-HYPERLINK-LIST", LegacyPptDiagnosticSeverity.Warning,
                    "An external-object list has an invalid record header; hyperlinks remain preserve-only.",
                    list.Offset);
                return;
            }
            LegacyPptRecord[] listAtoms = list.Children.Where(record =>
                record.Type == RecordExternalObjectListAtom).ToArray();
            if (listAtoms.Length != 1 || listAtoms[0].Version != 0
                || listAtoms[0].Instance != 0 || listAtoms[0].PayloadLength != 4) {
                AddDiagnostic("PPT-HYPERLINK-LIST-ATOM", LegacyPptDiagnosticSeverity.Warning,
                    "The external-object list has no unique valid ExObjListAtom; hyperlinks remain preserve-only.",
                    list.Offset);
                return;
            }

            uint greatestId = 0;
            foreach (LegacyPptRecord container in list.Children.Where(record =>
                         record.Type == RecordExternalHyperlink)) {
                LegacyPptHyperlink? hyperlink = TryReadHyperlink(container);
                if (hyperlink == null) continue;
                greatestId = Math.Max(greatestId, hyperlink.Id);
                if (_hyperlinksById.ContainsKey(hyperlink.Id)) {
                    AddDiagnostic("PPT-HYPERLINK-ID-DUPLICATE",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Hyperlink identifier {hyperlink.Id} occurs more than once; later targets remain preserve-only.",
                        container.Offset);
                    continue;
                }
                _hyperlinks.Add(hyperlink);
                _hyperlinksById.Add(hyperlink.Id, hyperlink);
            }

            uint seed = listAtoms[0].ReadUInt32(0);
            if (options.ReportUnsupportedContent && seed < greatestId) {
                AddDiagnostic("PPT-HYPERLINK-ID-SEED", LegacyPptDiagnosticSeverity.Warning,
                    $"The external-object id seed {seed} is below hyperlink id {greatestId}; new targets require a repaired seed.",
                    listAtoms[0].Offset);
            }
        }

        private static bool HasReadableExternalHyperlinkTarget(
            LegacyPptRecord container) => HasReadableExternalHyperlinkTarget(
            container, ReadRawChildHeaders(container));

        private static bool HasReadableExternalHyperlinkTarget(
            LegacyPptRecord source, RawRecordHeader container) =>
            HasReadableExternalHyperlinkTarget(source,
                ReadRawChildHeaders(source, container));

        private static bool HasReadableExternalHyperlinkTarget(
            LegacyPptRecord source,
            IEnumerable<RawRecordHeader> strings) {
            foreach (RawRecordHeader record in strings) {
                if (record.Type != RecordCString || record.Version != 0
                    || (record.Instance != 1 && record.Instance != 3)
                    || (record.PayloadLength & 1) != 0) continue;
                string value = source.ReadUtf16Text(record.PayloadOffset,
                    record.PayloadLength).TrimEnd('\0');
                if (string.IsNullOrEmpty(value)) continue;
                if (record.Instance == 1) return true;
                var hyperlink = new LegacyPptHyperlink(1,
                    friendlyName: null, target: null, location: value);
                if (!hyperlink.IsInternalSlideTarget) return true;
            }
            return false;
        }

        private LegacyPptHyperlink? TryReadHyperlink(LegacyPptRecord container) {
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-HYPERLINK-CONTAINER", LegacyPptDiagnosticSeverity.Warning,
                    "An ExHyperlinkContainer has an invalid record header and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordExternalHyperlinkAtom).ToArray();
            if (atoms.Length != 1 || atoms[0].Version != 0 || atoms[0].Instance != 0
                || atoms[0].PayloadLength != 4) {
                AddDiagnostic("PPT-HYPERLINK-ATOM", LegacyPptDiagnosticSeverity.Warning,
                    "An ExHyperlinkContainer has no unique valid ExHyperlinkAtom and remains preserve-only.",
                    container.Offset);
                return null;
            }
            uint id = atoms[0].ReadUInt32(0);
            if (id == 0) {
                AddDiagnostic("PPT-HYPERLINK-ID", LegacyPptDiagnosticSeverity.Warning,
                    "A hyperlink uses reserved identifier zero and remains preserve-only.", atoms[0].Offset);
                return null;
            }
            if (!TryReadHyperlinkString(container, 0, out string? friendlyName)
                || !TryReadHyperlinkString(container, 1, out string? target)
                || !TryReadHyperlinkString(container, 3, out string? location)) {
                AddDiagnostic("PPT-HYPERLINK-STRING", LegacyPptDiagnosticSeverity.Warning,
                    $"Hyperlink identifier {id} has duplicate or malformed text fields and remains preserve-only.",
                    container.Offset);
                return null;
            }
            if (string.IsNullOrEmpty(target) && string.IsNullOrEmpty(location)) {
                AddDiagnostic("PPT-HYPERLINK-TARGET", LegacyPptDiagnosticSeverity.Warning,
                    $"Hyperlink identifier {id} has no target or location and remains preserve-only.",
                    container.Offset);
                return null;
            }
            return new LegacyPptHyperlink(id, friendlyName, target, location);
        }

        private static bool TryReadHyperlinkString(LegacyPptRecord container,
            ushort instance, out string? value) {
            value = null;
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordCString && record.Instance == instance).ToArray();
            if (atoms.Length == 0) return true;
            if (atoms.Length != 1 || atoms[0].Version != 0
                || (atoms[0].PayloadLength & 1) != 0) return false;
            value = atoms[0].ReadUtf16Text().TrimEnd('\0');
            return true;
        }

        private IReadOnlyList<LegacyPptInteraction> ReadShapeInteractions(
            LegacyPptRecord shapeContainer, LegacyPptImportOptions options) {
            var interactions = new List<LegacyPptInteraction>();
            foreach (LegacyPptRecord clientData in shapeContainer.Children.Where(record =>
                         record.Type == OfficeArtClientData)) {
                if (HasRunProgramActionInOwner(clientData)) {
                    HasRunProgramContent = true;
                }
                foreach (LegacyPptRecord record in clientData.Children.Where(record =>
                             record.Type == RecordInteractiveInfo)) {
                    LegacyPptInteraction? interaction = TryReadInteraction(record, options);
                    if (interaction != null) interactions.Add(interaction);
                }
            }
            ReportDuplicateTriggers(interactions, shapeContainer.Offset, "shape");
            return new ReadOnlyCollection<LegacyPptInteraction>(interactions);
        }

        private IReadOnlyList<LegacyPptTextInteraction> ReadTextInteractions(
            LegacyPptRecord? textbox, int exposedTextLength,
            LegacyPptImportOptions options) {
            if (textbox == null) return Array.Empty<LegacyPptTextInteraction>();
            if (HasRunProgramActionInOwner(textbox)) {
                HasRunProgramContent = true;
            }
            var interactions = new List<LegacyPptTextInteraction>();
            IReadOnlyList<LegacyPptRecord> children = textbox.Children;
            for (int index = 0; index < children.Count; index++) {
                LegacyPptRecord actionRecord = children[index];
                if (actionRecord.Type != RecordInteractiveInfo) continue;
                if (HasRunProgramActionAtom(actionRecord)) {
                    HasRunProgramContent = true;
                }
                if (index + 1 >= children.Count
                    || children[index + 1].Type != RecordTextInteractiveInfoAtom) {
                    AddDiagnostic("PPT-TEXT-ACTION-RANGE-MISSING",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A text interaction has no following TextInteractiveInfoAtom and remains preserve-only.",
                        actionRecord.Offset);
                    continue;
                }
                LegacyPptRecord range = children[++index];
                LegacyPptInteraction? interaction = TryReadInteraction(actionRecord, options);
                if (interaction == null) continue;
                if (range.Version != 0 || range.Instance != (ushort)interaction.Trigger
                    || range.PayloadLength != 8) {
                    AddDiagnostic("PPT-TEXT-ACTION-RANGE", LegacyPptDiagnosticSeverity.Warning,
                        "A text interaction range has an invalid record header and remains preserve-only.",
                        range.Offset);
                    continue;
                }
                int begin = range.ReadInt32(0);
                int end = range.ReadInt32(4);
                int clippedEnd = Math.Min(end, exposedTextLength);
                if (begin < 0 || end <= begin || begin >= exposedTextLength
                    || clippedEnd <= begin) {
                    AddDiagnostic("PPT-TEXT-ACTION-BOUNDS", LegacyPptDiagnosticSeverity.Warning,
                        $"A text interaction range [{begin}, {end}) is outside the exposed text and remains preserve-only.",
                        range.Offset);
                    continue;
                }
                interactions.Add(new LegacyPptTextInteraction(begin,
                    clippedEnd - begin, interaction));
            }
            ReportOverlappingTextTriggers(interactions, textbox.Offset);
            return new ReadOnlyCollection<LegacyPptTextInteraction>(interactions);
        }

        private LegacyPptInteraction? TryReadInteraction(LegacyPptRecord container,
            LegacyPptImportOptions options) {
            if (HasRunProgramActionAtom(container)) {
                HasRunProgramContent = true;
            }
            if (container.Version != 0x0F || container.Instance > 1) {
                AddDiagnostic("PPT-ACTION-CONTAINER", LegacyPptDiagnosticSeverity.Warning,
                    "An InteractiveInfoContainer has an invalid trigger or record header and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordInteractiveInfoAtom).ToArray();
            if (atoms.Length != 1 || atoms[0].Version != 0 || atoms[0].Instance != 0
                || atoms[0].PayloadLength != 16) {
                AddDiagnostic("PPT-ACTION-ATOM", LegacyPptDiagnosticSeverity.Warning,
                    "An interaction has no unique valid InteractiveInfoAtom and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord atom = atoms[0];
            byte actionValue = atom.ReadByte(8);
            byte jumpValue = atom.ReadByte(10);
            byte hyperlinkTypeValue = atom.ReadByte(12);
            if (actionValue > (byte)LegacyPptInteractionAction.CustomShow
                || jumpValue > (byte)LegacyPptInteractionJump.EndShow
                || !IsDefinedHyperlinkType(hyperlinkTypeValue)) {
                AddDiagnostic("PPT-ACTION-ENUM", LegacyPptDiagnosticSeverity.Warning,
                    "An InteractiveInfoAtom contains an undefined action, jump, or hyperlink type and remains preserve-only.",
                    atom.Offset);
                return null;
            }
            byte flags = atom.ReadByte(11);
            if (options.ReportUnsupportedContent && (flags & 0xF0) != 0) {
                AddDiagnostic("PPT-ACTION-RESERVED", LegacyPptDiagnosticSeverity.Warning,
                    "An InteractiveInfoAtom has nonzero reserved flags that remain preserve-only.", atom.Offset);
            }
            uint hyperlinkId = atom.ReadUInt32(4);
            _hyperlinksById.TryGetValue(hyperlinkId, out LegacyPptHyperlink? hyperlink);
            LegacyPptInteractionAction action = (LegacyPptInteractionAction)actionValue;
            if (action == LegacyPptInteractionAction.Hyperlink && hyperlinkId != 0
                && hyperlink == null) {
                AddDiagnostic("PPT-ACTION-HYPERLINK-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"An interaction references missing hyperlink identifier {hyperlinkId} and remains preserve-only.",
                    atom.Offset);
            }
            string? name = null;
            LegacyPptRecord[] names = container.Children.Where(record =>
                record.Type == RecordCString && record.Instance == 0).ToArray();
            if (names.Length == 1 && names[0].Version == 0
                && (names[0].PayloadLength & 1) == 0) {
                name = names[0].ReadUtf16Text().TrimEnd('\0');
            } else if (names.Length > 0) {
                AddDiagnostic("PPT-ACTION-NAME", LegacyPptDiagnosticSeverity.Warning,
                    "An interaction has duplicate or malformed macro/program/show name data.",
                    container.Offset);
            }
            LegacyPptCustomShow? customShow = null;
            if (action == LegacyPptInteractionAction.CustomShow && name != null) {
                LegacyPptCustomShow[] matches = _customShows.Where(show =>
                    string.Equals(show.Name, name, StringComparison.Ordinal)).ToArray();
                if (matches.Length == 1 && matches[0].IsEditable) {
                    customShow = matches[0];
                }
            }
            var interaction = new LegacyPptInteraction(
                (LegacyPptInteractionTrigger)container.Instance, action,
                (LegacyPptInteractionJump)jumpValue,
                (LegacyPptHyperlinkType)hyperlinkTypeValue,
                atom.ReadUInt32(0), hyperlinkId, atom.ReadByte(9), flags,
                name, hyperlink, customShow);
            if (options.ReportUnsupportedContent && !IsNativelyProjectable(interaction)) {
                AddDiagnostic("PPT-ACTION-PRESERVE-ONLY",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The {interaction.Action} action cannot be represented by the current editable PowerPoint action surface and remains preserve-only.",
                    container.Offset);
            }
            return interaction;
        }

        private static bool HasRunProgramActionAtom(
            LegacyPptRecord container) => HasRunProgramActionAtom(container,
            ReadRawChildHeaders(container));

        private static bool HasRunProgramActionAtom(LegacyPptRecord source,
            IEnumerable<RawRecordHeader> records) => records.Any(record =>
            record.Type == RecordInteractiveInfoAtom
            && record.PayloadLength >= 9
            && source.ReadByte(record.PayloadOffset + 8) ==
            (byte)LegacyPptInteractionAction.RunProgram);

        private static bool HasRunProgramActionInOwner(
            LegacyPptRecord owner) => ReadRawChildHeaders(owner).Any(record =>
            record.Type == RecordInteractiveInfo
            && HasRunProgramActionAtom(owner,
                ReadRawChildHeaders(owner, record)));

        private void ReportDuplicateTriggers(IReadOnlyList<LegacyPptInteraction> interactions,
            long offset, string owner) {
            foreach (IGrouping<LegacyPptInteractionTrigger, LegacyPptInteraction> group
                     in interactions.GroupBy(interaction => interaction.Trigger)) {
                if (group.Count() < 2) continue;
                AddDiagnostic("PPT-ACTION-DUPLICATE", LegacyPptDiagnosticSeverity.Warning,
                    $"A {owner} contains multiple {group.Key} interactions; only one can be projected natively.",
                    offset);
            }
        }

        private void ReportOverlappingTextTriggers(
            IReadOnlyList<LegacyPptTextInteraction> interactions, long offset) {
            foreach (IGrouping<LegacyPptInteractionTrigger, LegacyPptTextInteraction> group
                     in interactions.GroupBy(item => item.Interaction.Trigger)) {
                int previousEnd = -1;
                foreach (LegacyPptTextInteraction item in group.OrderBy(item => item.Start)) {
                    if (item.Start < previousEnd) {
                        AddDiagnostic("PPT-TEXT-ACTION-OVERLAP",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"Text contains overlapping {group.Key} interaction ranges; the binary records remain preserve-only.",
                            offset);
                        break;
                    }
                    previousEnd = Math.Max(previousEnd,
                        checked(item.Start + item.Length));
                }
            }
        }

        private static bool IsDefinedHyperlinkType(byte value) => value <= 3
            || value == 6 || value == 7 || value == 8 || value == 9
            || value == 10 || value == 255;

        private bool IsNativelyProjectable(LegacyPptInteraction interaction) {
            byte allowedFlags = interaction.Action ==
                LegacyPptInteractionAction.CustomShow ? (byte)0x07 : (byte)0x03;
            if (interaction.OleVerb != 0
                || (interaction.Flags & ~allowedFlags) != 0) return false;
            if (interaction.SoundIdReference != 0) {
                LegacyPptSound? sound = FindSound(interaction.SoundIdReference);
                if (sound?.HasData != true || sound.ContentType == null) return false;
            }
            if (interaction.Action == LegacyPptInteractionAction.None) {
                return interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0
                    && string.IsNullOrEmpty(interaction.Name);
            }
            if (interaction.Action == LegacyPptInteractionAction.Macro) {
                return !string.IsNullOrEmpty(interaction.Name)
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0;
            }
            if (interaction.Action == LegacyPptInteractionAction.RunProgram) {
                return !string.IsNullOrEmpty(interaction.Name)
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0
                    && Uri.TryCreate(interaction.Name, UriKind.RelativeOrAbsolute,
                        out _);
            }
            if (interaction.Action == LegacyPptInteractionAction.CustomShow) {
                return interaction.CustomShow != null
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0;
            }
            if (interaction.Action == LegacyPptInteractionAction.Jump) {
                return interaction.Jump != LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0
                    && string.IsNullOrEmpty(interaction.Name);
            }
            if (interaction.Action != LegacyPptInteractionAction.Hyperlink) return false;
            if (interaction.Jump != LegacyPptInteractionJump.None
                || !string.IsNullOrEmpty(interaction.Name)
                || interaction.Hyperlink?.ExtensionFlags != 0
                || interaction.HyperlinkType == LegacyPptHyperlinkType.CustomShow) {
                return false;
            }
            return (interaction.HyperlinkType != LegacyPptHyperlinkType.SlideNumber
                    && interaction.Hyperlink?.Uri != null)
                || (interaction.HyperlinkType == LegacyPptHyperlinkType.SlideNumber
                    && interaction.Hyperlink?.IsInternalSlideTarget == true)
                || interaction.HyperlinkType == LegacyPptHyperlinkType.NextSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.PreviousSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.FirstSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.LastSlide;
        }
    }
}
