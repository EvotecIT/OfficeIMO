using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordDocumentAtom = 0x03E9;
        private const ushort RecordExternalObjectList = 0x0409;
        private const ushort RecordExternalObjectListAtom = 0x040A;
        private const ushort RecordExternalHyperlink = 0x0FD7;
        private const ushort RecordExternalHyperlinkAtom = 0x0FD3;
        private const ushort RecordInteractiveInfo = 0x0FF2;
        private const ushort RecordInteractiveInfoAtom = 0x0FF3;
        private const ushort RecordTextInteractiveInfoAtom = 0x0FDF;
        private const ushort OfficeArtClientData = 0xF011;

        private static bool ShapeInteractionsEqual(LegacyPptShapeProjection source,
            LegacyPptWriter.LegacyPptWriterShapeInteractions current,
            LegacyPptWriter.LegacyPptWriterInteractionCatalog catalog,
            LegacyPptProjectionMap projectionMap) =>
            InteractionListsEqual(source.ShapeInteractions,
                current.ShapeInteractions, catalog, projectionMap);

        private static bool TextInteractionsEqual(LegacyPptShapeProjection source,
            LegacyPptWriter.LegacyPptWriterShapeInteractions current,
            LegacyPptWriter.LegacyPptWriterInteractionCatalog catalog,
            LegacyPptProjectionMap projectionMap) {
            if (source.TextInteractions.Count != current.TextInteractions.Count) return false;
            for (int index = 0; index < source.TextInteractions.Count; index++) {
                LegacyPptTextInteraction left = source.TextInteractions[index];
                LegacyPptWriter.LegacyPptWriterTextInteraction right =
                    current.TextInteractions[index];
                if (left.Start != right.Begin
                    || left.Start + left.Length != right.End
                    || !InteractionsEqual(left.Interaction, right.Interaction,
                        catalog, projectionMap)) return false;
            }
            return true;
        }

        private static bool InteractionListsEqual(
            IReadOnlyList<LegacyPptInteraction> source,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterInteraction> current,
            LegacyPptWriter.LegacyPptWriterInteractionCatalog catalog,
            LegacyPptProjectionMap projectionMap) {
            if (source.Count != current.Count) return false;
            for (int index = 0; index < source.Count; index++) {
                if (!InteractionsEqual(source[index], current[index], catalog,
                        projectionMap)) return false;
            }
            return true;
        }

        private static bool InteractionsEqual(LegacyPptInteraction source,
            LegacyPptWriter.LegacyPptWriterInteraction current,
            LegacyPptWriter.LegacyPptWriterInteractionCatalog catalog,
            LegacyPptProjectionMap projectionMap) {
            LegacyPptInteractionAction action = source.Action;
            LegacyPptInteractionJump jump = source.Jump;
            if (action == LegacyPptInteractionAction.Hyperlink
                && TryMapHyperlinkTypeToJump(source.HyperlinkType,
                    out LegacyPptInteractionJump hyperlinkJump)) {
                action = LegacyPptInteractionAction.Jump;
                jump = hyperlinkJump;
            }
            if (source.Trigger != current.Trigger || action != current.Action
                || jump != current.Jump
                || source.SoundIdReference != current.SoundIdReference
                || source.OleVerb != current.OleVerb
                || source.Flags != current.Flags) return false;
            if (action != LegacyPptInteractionAction.Hyperlink) {
                return action is LegacyPptInteractionAction.Macro
                        or LegacyPptInteractionAction.RunProgram
                        or LegacyPptInteractionAction.CustomShow
                    ? string.Equals(source.Name, current.Name,
                        StringComparison.Ordinal)
                    : string.IsNullOrEmpty(source.Name)
                        && string.IsNullOrEmpty(current.Name);
            }
            if (!string.IsNullOrEmpty(source.Name)
                || !string.IsNullOrEmpty(current.Name)) return false;
            LegacyPptWriter.LegacyPptWriterHyperlink? currentHyperlink =
                catalog.FindHyperlink(current.HyperlinkIdReference);
            if (source.Hyperlink?.TargetSlideId is uint sourceSlideId) {
                return projectionMap.TryGetSlide(sourceSlideId,
                           out LegacyPptSlideProjection? targetSlide)
                    && targetSlide != null
                    && string.Equals(targetSlide.SlidePartUri,
                        currentHyperlink?.TargetSlidePartUri,
                        StringComparison.Ordinal)
                    && string.Equals(source.Hyperlink.ScreenTip,
                        currentHyperlink?.ScreenTip, StringComparison.Ordinal);
            }
            string? sourceTarget = source.Hyperlink?.Uri?.OriginalString;
            string? currentTarget = currentHyperlink?.Target;
            return string.Equals(sourceTarget, currentTarget,
                    StringComparison.Ordinal)
                && string.Equals(source.Hyperlink?.ScreenTip,
                    currentHyperlink?.ScreenTip, StringComparison.Ordinal);
        }

        private static bool TryMapHyperlinkTypeToJump(LegacyPptHyperlinkType type,
            out LegacyPptInteractionJump jump) {
            jump = type switch {
                LegacyPptHyperlinkType.NextSlide => LegacyPptInteractionJump.NextSlide,
                LegacyPptHyperlinkType.PreviousSlide => LegacyPptInteractionJump.PreviousSlide,
                LegacyPptHyperlinkType.FirstSlide => LegacyPptInteractionJump.FirstSlide,
                LegacyPptHyperlinkType.LastSlide => LegacyPptInteractionJump.LastSlide,
                _ => LegacyPptInteractionJump.None
            };
            return jump != LegacyPptInteractionJump.None;
        }

        private static bool TryCreateInteractionContext(PowerPointPresentation presentation,
            LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap,
            LegacyPptWriter.LegacyPptWriterInteractionCatalog catalog,
            out PreservingInteractionContext context) {
            if (!TryReadExternalObjectIdSeed(package, out uint externalObjectIdSeed)) {
                context = null!;
                return false;
            }
            context = new PreservingInteractionContext(presentation, projectionMap,
                externalObjectIdSeed);
            foreach (LegacyPptWriter.LegacyPptWriterHyperlink hyperlink
                     in catalog.Hyperlinks) {
                if (!context.TryMap(hyperlink, out _)) return false;
            }
            return true;
        }

        private static bool TryReadExternalObjectIdSeed(
            LegacyPptPackage package,
            out uint seed) {
            seed = 0;
            if (!TryReadDocument(package, out LegacyPptRecord? document)
                || document == null) return true;
            LegacyPptRecord[] lists = document.Children.Where(record =>
                record.Type == RecordExternalObjectList).Take(2).ToArray();
            if (lists.Length > 1) return false;
            if (lists.Length == 0) return true;
            LegacyPptRecord[] atoms = lists[0].Children.Where(record =>
                record.Type == RecordExternalObjectListAtom).Take(2).ToArray();
            if (atoms.Length > 1) return false;
            if (atoms.Length == 0 || atoms[0].PayloadLength != 4) return true;
            seed = atoms[0].ReadUInt32(0);
            return true;
        }

        private static bool TryAppendNewHyperlinks(LegacyPptPackage package,
            byte[]? currentDocumentBytes,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterHyperlink> hyperlinks,
            out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (hyperlinks.Count == 0) return false;
            LegacyPptRecord document;
            if (currentDocumentBytes != null) {
                document = LegacyPptRecordReader.ReadSingle(currentDocumentBytes, 0,
                    new LegacyPptImportOptions());
            } else if (!TryReadDocument(package, out LegacyPptRecord? source)
                       || source == null) {
                return false;
            } else {
                document = source;
            }

            LegacyPptRecord[] lists = document.Children.Where(record =>
                record.Type == RecordExternalObjectList).ToArray();
            if (lists.Length > 1) return false;
            byte[] rewrittenList;
            if (lists.Length == 0) {
                var seedPayload = new byte[4];
                WriteUInt32(seedPayload, 0, hyperlinks.Max(link => link.Id));
                var listChildren = new List<byte[]> {
                    BuildRecord(version: 0, instance: 0,
                        RecordExternalObjectListAtom, seedPayload)
                };
                listChildren.AddRange(hyperlinks.Select(link =>
                    LegacyPptWriter.BuildExternalHyperlinkRecord(link)));
                rewrittenList = BuildRecord(version: 0x0F, instance: 0,
                    RecordExternalObjectList, Concat(listChildren));
            } else {
                LegacyPptRecord list = lists[0];
                LegacyPptRecord[] atoms = list.Children.Where(record =>
                    record.Type == RecordExternalObjectListAtom).ToArray();
                if (list.Version != 0x0F || list.Instance != 0 || atoms.Length != 1
                    || atoms[0].Version != 0 || atoms[0].Instance != 0
                    || atoms[0].PayloadLength != 4) return false;
                var listChildren = new List<byte[]>(list.Children.Count + hyperlinks.Count);
                foreach (LegacyPptRecord child in list.Children) {
                    if (ReferenceEquals(child, atoms[0])) {
                        byte[] atom = child.CopyRecordBytes();
                        WriteUInt32(atom, 8, Math.Max(child.ReadUInt32(0),
                            hyperlinks.Max(link => link.Id)));
                        listChildren.Add(atom);
                    } else {
                        listChildren.Add(child.CopyRecordBytes());
                    }
                }
                listChildren.AddRange(hyperlinks.Select(link =>
                    LegacyPptWriter.BuildExternalHyperlinkRecord(link)));
                rewrittenList = BuildRecord(list.Version, list.Instance,
                    list.Type, Concat(listChildren));
            }

            var documentChildren = new List<byte[]>(document.Children.Count + 1);
            bool inserted = false;
            foreach (LegacyPptRecord child in document.Children) {
                if (lists.Length == 1 && ReferenceEquals(child, lists[0])) {
                    documentChildren.Add(rewrittenList);
                    inserted = true;
                } else {
                    documentChildren.Add(child.CopyRecordBytes());
                    if (lists.Length == 0 && child.Type == RecordDocumentAtom) {
                        documentChildren.Add(rewrittenList);
                        inserted = true;
                    }
                }
            }
            if (!inserted) return false;
            byte[] rebuilt = BuildRecord(document.Version, document.Instance,
                document.Type, Concat(documentChildren));
            LegacyPptRecord rebuiltRecord = LegacyPptRecordReader.ReadSingle(rebuilt, 0,
                new LegacyPptImportOptions());
            return LegacyPptWriter.TryRewriteDocumentHyperlinkExtensions(
                rebuiltRecord, hyperlinks, replaceExisting: false, out bytes);
        }

        private static bool TryRewriteClientDataInteractions(LegacyPptRecord clientData,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterInteraction> interactions,
            bool append, out byte[] bytes) {
            if (clientData.Version != 0x0F) {
                bytes = clientData.CopyRecordBytes();
                return false;
            }
            var children = new List<byte[]>(clientData.Children.Count + interactions.Count);
            foreach (LegacyPptRecord child in clientData.Children) {
                if (child.Type == RecordInteractiveInfo) {
                    if (!IsRewritableInteractiveInfo(child)) {
                        bytes = clientData.CopyRecordBytes();
                        return false;
                    }
                    continue;
                }
                children.Add(child.CopyRecordBytes());
            }
            if (append) {
                children.AddRange(interactions.Select(
                    LegacyPptWriter.BuildInteractiveInfoRecord));
            }
            bytes = BuildRecord(clientData.Version, clientData.Instance,
                clientData.Type, Concat(children));
            return true;
        }

        private static bool TryRewriteTextInteractions(LegacyPptRecord textbox,
            string originalText, string? replacementText,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterTextInteraction> interactions,
            out byte[] bytes) {
            if (textbox.Version != 0x0F) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }
            LegacyPptRecord[] textRecords = textbox.Children.Where(record =>
                record.Type == RecordTextChars || record.Type == RecordTextBytes).ToArray();
            if (replacementText != null && textRecords.Length != 1) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }
            byte[]? replacementRecord = null;
            if (replacementText != null && !TryBuildTextRecord(textbox, textRecords[0],
                    originalText, replacementText, out replacementRecord)) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }

            var children = new List<byte[]>(textbox.Children.Count + interactions.Count * 2);
            for (int index = 0; index < textbox.Children.Count; index++) {
                LegacyPptRecord child = textbox.Children[index];
                if (child.Type == RecordInteractiveInfo) {
                    if (!IsRewritableInteractiveInfo(child)
                        || index + 1 >= textbox.Children.Count
                        || textbox.Children[index + 1].Type != RecordTextInteractiveInfoAtom
                        || textbox.Children[index + 1].Version != 0
                        || textbox.Children[index + 1].Instance != child.Instance
                        || textbox.Children[index + 1].PayloadLength != 8) {
                        bytes = textbox.CopyRecordBytes();
                        return false;
                    }
                    index++;
                    continue;
                }
                if (child.Type == RecordTextInteractiveInfoAtom) {
                    bytes = textbox.CopyRecordBytes();
                    return false;
                }
                children.Add(replacementRecord != null && ReferenceEquals(child, textRecords[0])
                    ? replacementRecord
                    : child.CopyRecordBytes());
            }
            foreach (LegacyPptWriter.LegacyPptWriterTextInteraction interaction
                     in interactions) {
                children.Add(LegacyPptWriter.BuildInteractiveInfoRecord(
                    interaction.Interaction));
                children.Add(LegacyPptWriter.BuildTextInteractiveInfoRecord(interaction));
            }
            bytes = BuildRecord(textbox.Version, textbox.Instance, textbox.Type,
                Concat(children));
            return true;
        }

        private static bool IsRewritableInteractiveInfo(LegacyPptRecord container) {
            if (container.Version != 0x0F || container.Instance > 1) return false;
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordInteractiveInfoAtom).ToArray();
            return atoms.Length == 1 && atoms[0].Version == 0
                && atoms[0].Instance == 0 && atoms[0].PayloadLength == 16
                && container.Children.All(child => child.Type == RecordInteractiveInfoAtom
                    || child.Type == RecordCString);
        }

        private sealed class PreservingInteractionContext {
            private readonly Dictionary<string, uint> _idsByTarget =
                new(StringComparer.Ordinal);
            private readonly Dictionary<uint, uint> _binaryIdsByCatalogId = new();
            private readonly List<LegacyPptWriter.LegacyPptWriterHyperlink> _newHyperlinks = new();
            private readonly Dictionary<string, LegacyPptWriter.LegacyPptWriterSlideTarget>
                _slideTargetsByPartUri = new(StringComparer.Ordinal);
            private uint _nextId;

            internal PreservingInteractionContext(
                PowerPointPresentation presentation,
                LegacyPptProjectionMap projectionMap, uint objectIdSeed) {
                _nextId = objectIdSeed;
                uint nextSlideId = projectionMap.Slides.Count == 0
                    ? 255U
                    : projectionMap.Slides.Max(slide => slide.SlideId);
                for (int index = 0; index < presentation.Slides.Count; index++) {
                    PowerPointSlide slide = presentation.Slides[index];
                    uint binarySlideId;
                    if (projectionMap.TryGetSlide(slide,
                            out LegacyPptSlideProjection? sourceSlide)
                        && sourceSlide != null) {
                        binarySlideId = sourceSlide.SlideId;
                    } else {
                        if (nextSlideId >= 0x7FFFFFFFU) continue;
                        binarySlideId = ++nextSlideId;
                    }
                    string partUri = slide.SlidePart.Uri.ToString();
                    _slideTargetsByPartUri[partUri] =
                        new LegacyPptWriter.LegacyPptWriterSlideTarget(
                            partUri, binarySlideId, index + 1,
                            slide.SlidePart.Slide?.CommonSlideData?.Name?.Value);
                }
                foreach (LegacyPptHyperlink hyperlink in projectionMap.Hyperlinks) {
                    _nextId = Math.Max(_nextId, hyperlink.Id);
                    string? key;
                    if (hyperlink.TargetSlideId is uint targetSlideId
                        && projectionMap.TryGetSlide(targetSlideId,
                            out LegacyPptSlideProjection? targetSlide)
                        && targetSlide != null) {
                        key = LegacyPptWriter.CreateInternalHyperlinkKey(
                            targetSlide.SlidePartUri, hyperlink.ScreenTip);
                    } else {
                        string? target = hyperlink.Uri?.OriginalString;
                        key = target == null ? null :
                            LegacyPptWriter.CreateHyperlinkKey(target,
                                hyperlink.ScreenTip);
                    }
                    if (key != null && !_idsByTarget.ContainsKey(key)) {
                        _idsByTarget.Add(key, hyperlink.Id);
                    }
                }
            }

            internal IReadOnlyList<LegacyPptWriter.LegacyPptWriterHyperlink> NewHyperlinks =>
                _newHyperlinks;

            internal uint? ResolveSlideId(string partUri) =>
                _slideTargetsByPartUri.TryGetValue(partUri,
                    out LegacyPptWriter.LegacyPptWriterSlideTarget target)
                    ? target.BinarySlideId
                    : (uint?)null;

            internal bool TryMap(LegacyPptWriter.LegacyPptWriterHyperlink hyperlink,
                out uint binaryId) {
                if (_binaryIdsByCatalogId.TryGetValue(hyperlink.Id, out binaryId)) return true;
                string key;
                LegacyPptWriter.LegacyPptWriterSlideTarget slideTarget = default;
                if (hyperlink.TargetSlidePartUri != null) {
                    if (!_slideTargetsByPartUri.TryGetValue(
                            hyperlink.TargetSlidePartUri, out slideTarget)) return false;
                    key = LegacyPptWriter.CreateInternalHyperlinkKey(
                        hyperlink.TargetSlidePartUri, hyperlink.ScreenTip);
                } else {
                    if (hyperlink.Target == null) return false;
                    key = LegacyPptWriter.CreateHyperlinkKey(hyperlink.Target,
                        hyperlink.ScreenTip);
                }
                if (!_idsByTarget.TryGetValue(key, out binaryId)) {
                    if (_nextId == uint.MaxValue) return false;
                    binaryId = ++_nextId;
                    _idsByTarget.Add(key, binaryId);
                    _newHyperlinks.Add(hyperlink.TargetSlidePartUri != null
                        ? new LegacyPptWriter.LegacyPptWriterHyperlink(
                            binaryId, slideTarget.PartUri,
                            slideTarget.BinarySlideId, slideTarget.SlideNumber,
                            slideTarget.Name, hyperlink.ScreenTip,
                            hyperlink.ExtensionFlags)
                        : new LegacyPptWriter.LegacyPptWriterHyperlink(
                            binaryId, hyperlink.Target!, hyperlink.ScreenTip,
                            hyperlink.ExtensionFlags));
                }
                _binaryIdsByCatalogId.Add(hyperlink.Id, binaryId);
                return true;
            }

            internal LegacyPptWriter.LegacyPptWriterShapeInteractions Remap(
                LegacyPptWriter.LegacyPptWriterShapeInteractions source) => new(
                    source.ShapeInteractions.Select(Remap).ToArray(),
                    source.TextInteractions.Select(item =>
                        new LegacyPptWriter.LegacyPptWriterTextInteraction(
                            item.Begin, item.End, Remap(item.Interaction))).ToArray());

            internal LegacyPptWriter.LegacyPptWriterInteractionCatalog RemapCatalog(
                IEnumerable<PowerPointSlide> slides,
                LegacyPptWriter.LegacyPptWriterInteractionCatalog source) {
                var result = new LegacyPptWriter.LegacyPptWriterInteractionCatalog();
                foreach (PowerPointShape shape in slides.SelectMany(slide => slide.Shapes)) {
                    LegacyPptWriter.LegacyPptWriterShapeInteractions interactions =
                        source.Get(shape);
                    if (interactions.HasInteractions) {
                        result.Add(shape.Element, Remap(interactions));
                    }
                }
                return result;
            }

            private LegacyPptWriter.LegacyPptWriterInteraction Remap(
                LegacyPptWriter.LegacyPptWriterInteraction source) {
                uint hyperlinkId = source.HyperlinkIdReference == 0
                    ? 0
                    : _binaryIdsByCatalogId[source.HyperlinkIdReference];
                return new LegacyPptWriter.LegacyPptWriterInteraction(source.Trigger,
                    source.Action, source.Jump, source.HyperlinkType, hyperlinkId,
                    source.Name, source.SoundIdReference, source.OleVerb,
                    source.Flags);
            }
        }

        private sealed class ProjectedInteractionEdit {
            internal ProjectedInteractionEdit(
                LegacyPptWriter.LegacyPptWriterShapeInteractions interactions,
                bool rewriteShapeInteractions, bool rewriteTextInteractions) {
                Interactions = interactions;
                RewriteShapeInteractions = rewriteShapeInteractions;
                RewriteTextInteractions = rewriteTextInteractions;
            }

            internal LegacyPptWriter.LegacyPptWriterShapeInteractions Interactions { get; }
            internal bool RewriteShapeInteractions { get; }
            internal bool RewriteTextInteractions { get; }
        }
    }
}
