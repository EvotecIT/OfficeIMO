using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal sealed partial class WordDocumentMerge {
        private readonly Dictionary<int, int> _numberingMap = new();

        /// <summary>
        /// Copies the source list definitions. Abstract definitions are inserted before the destination's
        /// numbering instances to keep the schema order (abstractNum*, num*, numIdMacAtCleanup?).
        /// </summary>
        private void ImportNumbering() {
            Numbering? sourceNumbering = _sourceMain.NumberingDefinitionsPart?.Numbering;
            if (sourceNumbering == null || !sourceNumbering.Elements<AbstractNum>().Any() && !sourceNumbering.Elements<NumberingInstance>().Any()) {
                return;
            }

            Numbering destinationNumbering = EnsureDestinationNumbering();
            var abstractMap = new Dictionary<int, int>();
            var usedNsids = new HashSet<string>(
                destinationNumbering.Elements<AbstractNum>().Select(_ => _.Nsid?.Val?.Value).OfType<string>(),
                StringComparer.OrdinalIgnoreCase);

            int nextAbstractId = destinationNumbering.Elements<AbstractNum>()
                .Select(_ => _.AbstractNumberId?.Value ?? -1)
                .DefaultIfEmpty(-1)
                .Max() + 1;
            OpenXmlElement? abstractAnchor = destinationNumbering.Elements<NumberingInstance>().FirstOrDefault() ??
                (OpenXmlElement?)destinationNumbering.Elements<NumberingIdMacAtCleanup>().FirstOrDefault();

            foreach (AbstractNum abstractNum in sourceNumbering.Elements<AbstractNum>()) {
                int? oldId = abstractNum.AbstractNumberId?.Value;
                if (oldId == null) {
                    continue;
                }

                var clone = (AbstractNum)abstractNum.CloneNode(true);
                clone.AbstractNumberId = nextAbstractId;
                abstractMap[oldId.Value] = nextAbstractId++;

                string? nsid = clone.Nsid?.Val?.Value;
                if (nsid != null && !usedNsids.Add(nsid)) {
                    clone.Nsid!.Val = CreateUniqueNsid(usedNsids);
                }

                // Picture bullets live in the source numbering part and are not imported; fall back to the text bullet.
                foreach (LevelPictureBulletId pictureBullet in clone.Descendants<LevelPictureBulletId>().ToList()) {
                    pictureBullet.Remove();
                }

                RemapStyleReferences(clone);
                if (abstractAnchor != null) {
                    destinationNumbering.InsertBefore(clone, abstractAnchor);
                } else {
                    destinationNumbering.Append(clone);
                }
            }

            int nextNumberId = destinationNumbering.Elements<NumberingInstance>()
                .Select(_ => _.NumberID?.Value ?? 0)
                .DefaultIfEmpty(0)
                .Max() + 1;
            OpenXmlElement? numberAnchor = destinationNumbering.Elements<NumberingIdMacAtCleanup>().FirstOrDefault();

            foreach (NumberingInstance instance in sourceNumbering.Elements<NumberingInstance>()) {
                int? oldId = instance.NumberID?.Value;
                if (oldId == null) {
                    continue;
                }

                var clone = (NumberingInstance)instance.CloneNode(true);
                clone.NumberID = nextNumberId;
                _numberingMap[oldId.Value] = nextNumberId++;

                AbstractNumId? abstractNumId = clone.GetFirstChild<AbstractNumId>();
                if (abstractNumId?.Val?.Value is int abstractId && abstractMap.TryGetValue(abstractId, out int mappedAbstract)) {
                    abstractNumId.Val = mappedAbstract;
                }

                if (numberAnchor != null) {
                    destinationNumbering.InsertBefore(clone, numberAnchor);
                } else {
                    destinationNumbering.Append(clone);
                }
            }
        }

        private void RemapNumberingReferences(OpenXmlElement root) {
            if (_numberingMap.Count == 0) {
                return;
            }

            foreach (NumberingId numberingId in root.Descendants<NumberingId>()) {
                if (numberingId.Val?.Value is int oldId && _numberingMap.TryGetValue(oldId, out int mapped)) {
                    numberingId.Val = mapped;
                }
            }
        }

        private Numbering EnsureDestinationNumbering() {
            NumberingDefinitionsPart part = _destinationMain.NumberingDefinitionsPart ?? _destinationMain.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering ??= new Numbering();
            return part.Numbering;
        }

        private static string CreateUniqueNsid(HashSet<string> usedNsids) {
            uint next = usedNsids
                .Select(_ => uint.TryParse(_, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint parsed) ? parsed : 0U)
                .DefaultIfEmpty(0U)
                .Max();
            string candidate;
            do {
                next = next == uint.MaxValue ? 1U : next + 1U;
                candidate = next.ToString("X8", CultureInfo.InvariantCulture);
            } while (!usedNsids.Add(candidate));

            return candidate;
        }
    }
}
