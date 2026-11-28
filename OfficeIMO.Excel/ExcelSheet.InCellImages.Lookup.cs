using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.OpenXml.Internal;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;
using RichRel = DocumentFormat.OpenXml.Office.Y2022.RichValueRel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class InCellImageLookup {
            private readonly Metadata _metadata;
            private readonly List<MetadataBlock> _metadataBlocks;
            private readonly List<FutureMetadataBlock> _futureBlocks;
            private readonly List<Rich.RichValue> _richValues;
            private readonly List<Rich.RichValueStructure> _structures;
            private readonly List<RichRel.RichValueRelRelationship> _relationships;
            private readonly ExtendedPart _relationshipPart;
            private readonly RichImageIndex _index;
            private readonly Dictionary<uint, (OpenXmlPart Part, string AltText)> _resolved =
                new Dictionary<uint, (OpenXmlPart Part, string AltText)>();
            private readonly HashSet<uint> _invalid = new HashSet<uint>();

            internal InCellImageLookup(
                Metadata metadata,
                List<MetadataBlock> metadataBlocks,
                List<FutureMetadataBlock> futureBlocks,
                List<Rich.RichValue> richValues,
                List<Rich.RichValueStructure> structures,
                List<RichRel.RichValueRelRelationship> relationships,
                ExtendedPart relationshipPart) {
                _metadata = metadata;
                _metadataBlocks = metadataBlocks;
                _futureBlocks = futureBlocks;
                _richValues = richValues;
                _structures = structures;
                _relationships = relationships;
                _relationshipPart = relationshipPart;
                _index = new RichImageIndex(OfficeXmlNode.FromSdk(metadata),
                    index => OfficeXmlNode.FromSdk(metadataBlocks.ElementAtOrDefault(index)),
                    index => OfficeXmlNode.FromSdk(futureBlocks.ElementAtOrDefault(index)),
                    index => OfficeXmlNode.FromSdk(richValues.ElementAtOrDefault(index)),
                    index => OfficeXmlNode.FromSdk(structures.ElementAtOrDefault(index)));
            }

            internal Metadata Metadata => _metadata;
            internal IReadOnlyList<MetadataBlock> MetadataBlocks => _metadataBlocks;
            internal IReadOnlyList<FutureMetadataBlock> FutureBlocks => _futureBlocks;
            internal IReadOnlyList<Rich.RichValue> RichValues => _richValues;
            internal IReadOnlyList<Rich.RichValueStructure> Structures => _structures;
            internal IReadOnlyList<RichRel.RichValueRelRelationship> Relationships => _relationships;
            internal ExtendedPart RelationshipPart => _relationshipPart;

            internal bool TryResolve(Cell cell, out OpenXmlPart? imagePart, out string altText) {
                imagePart = null;
                altText = string.Empty;
                uint metadataIndex = cell.ValueMetaIndex?.Value ?? 0U;
                if (metadataIndex == 0U || _invalid.Contains(metadataIndex)) return false;
                if (_resolved.TryGetValue(metadataIndex, out (OpenXmlPart Part, string AltText) cached)) {
                    imagePart = cached.Part;
                    altText = cached.AltText;
                    return true;
                }
                if (!TryResolveSlot(metadataIndex, out InCellImageSlot slot)) {
                    _invalid.Add(metadataIndex);
                    return false;
                }
                RichRel.RichValueRelRelationship relationship = _relationships[slot.RelationshipIndex];
                if (relationship.Id?.Value is not string relationshipId) {
                    _invalid.Add(metadataIndex);
                    return false;
                }
                try {
                    imagePart = _relationshipPart.GetPartById(relationshipId);
                } catch (ArgumentOutOfRangeException) {
                    _invalid.Add(metadataIndex);
                    return false;
                }
                if (imagePart == null || !imagePart.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) {
                    _invalid.Add(metadataIndex);
                    imagePart = null;
                    return false;
                }
                altText = slot.AltText;
                _resolved[metadataIndex] = (imagePart, altText);
                return true;
            }

            internal bool TryResolveSlot(uint metadataIndex, out InCellImageSlot slot) {
                slot = default;
                if (!_index.TryResolve(metadataIndex, out var image) || image.Relationship >= _relationships.Count) return false;
                slot = new InCellImageSlot(metadataIndex, image.Future, image.Value, image.Relationship,
                    _richValues[image.Value], _structures[image.Structure], image.AltText);
                return true;
            }
        }

        private readonly struct InCellImageSlot {
            internal InCellImageSlot(
                uint metadataIndex,
                int futureIndex,
                int valueIndex,
                int relationshipIndex,
                Rich.RichValue value,
                Rich.RichValueStructure structure,
                string altText) {
                MetadataIndex = metadataIndex;
                FutureIndex = futureIndex;
                ValueIndex = valueIndex;
                RelationshipIndex = relationshipIndex;
                Value = value;
                Structure = structure;
                AltText = altText;
            }

            internal uint MetadataIndex { get; }
            internal int FutureIndex { get; }
            internal int ValueIndex { get; }
            internal int RelationshipIndex { get; }
            internal Rich.RichValue Value { get; }
            internal Rich.RichValueStructure Structure { get; }
            internal string AltText { get; }
        }

        private bool TryCreateInCellImageLookup(out InCellImageLookup? lookup) {
            lookup = null;
            WorkbookPart workbookPart = _excelDocument.WorkbookPartRoot;
            CellMetadataPart? metadataPart = workbookPart.CellMetadataPart;
            RdRichValuePart? valuePart = workbookPart.RdRichValueParts.FirstOrDefault();
            RdRichValueStructurePart? structurePart = workbookPart.GetPartsOfType<RdRichValueStructurePart>()
                .FirstOrDefault();
            ExtendedPart? relationshipPart = workbookPart.Parts.Select(pair => pair.OpenXmlPart).OfType<ExtendedPart>()
                .FirstOrDefault(part => string.Equals(part.RelationshipType, RichValueRelRelationshipType, StringComparison.Ordinal));
            if (metadataPart == null || valuePart == null || structurePart == null || relationshipPart == null) {
                return false;
            }

            ValidateInCellImageMetadataPart(metadataPart, "Cell metadata");
            ValidateInCellImageMetadataPart(valuePart, "Rich-value data");
            ValidateInCellImageMetadataPart(structurePart, "Rich-value structures");

            Metadata? metadata = metadataPart.Metadata;
            ValueMetadata? valueMetadata = metadata?.GetFirstChild<ValueMetadata>();
            FutureMetadata? future = metadata?.Elements<FutureMetadata>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, RichValueMetadataName, StringComparison.OrdinalIgnoreCase));
            Rich.RichValueData? richValueData = valuePart.RichValueData;
            Rich.RichValueStructures? structures = structurePart.RichValueStructures;
            if (metadata == null
                || valueMetadata == null
                || future == null
                || richValueData == null
                || structures == null) {
                return false;
            }
            lookup = new InCellImageLookup(
                metadata,
                valueMetadata.Elements<MetadataBlock>().ToList(),
                future.Elements<FutureMetadataBlock>().ToList(),
                richValueData.Elements<Rich.RichValue>().ToList(),
                structures.Elements<Rich.RichValueStructure>().ToList(),
                LoadRichValueRelationships(relationshipPart).Elements<RichRel.RichValueRelRelationship>().ToList(),
                relationshipPart);
            return true;
        }

        private bool TryReplaceExclusiveInCellImage(
            Cell cell,
            Stream imageStream,
            string contentType,
            string altText,
            out OpenXmlPart? resolvedImagePart) {
            resolvedImagePart = null;
            if (!TryCreateInCellImageLookup(out InCellImageLookup? lookup)
                || !lookup!.TryResolveSlot(cell.ValueMetaIndex?.Value ?? 0U, out InCellImageSlot slot)
                || !IsExclusiveInCellImageSlot(lookup, slot)) {
                return false;
            }
            RichRel.RichValueRelRelationship relationship = lookup.Relationships[slot.RelationshipIndex];
            if (relationship.Id?.Value is not string relationshipId) return false;
            OpenXmlPart currentImagePart;
            try {
                currentImagePart = lookup.RelationshipPart.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return false;
            }
            if (!currentImagePart.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return false;

            if (imageStream.CanSeek) imageStream.Position = 0;
            if (string.Equals(currentImagePart.ContentType, contentType, StringComparison.OrdinalIgnoreCase)) {
                currentImagePart.FeedData(imageStream);
                resolvedImagePart = currentImagePart;
            } else {
                ExtendedPart replacement = lookup.RelationshipPart.AddExtendedPart(
                    ImageRelationshipType,
                    contentType,
                    GetImageExtension(contentType));
                replacement.FeedData(imageStream);
                relationship.Id = lookup.RelationshipPart.GetIdOfPart(replacement);
                var relationships = new RichRel.RichValueRels(lookup.Relationships.Select(item => item.CloneNode(true)));
                SaveExtendedRoot(lookup.RelationshipPart, relationships);
                lookup.RelationshipPart.DeletePart(currentImagePart);
                resolvedImagePart = replacement;
            }
            SetRichValueAltText(slot.Value, slot.Structure, altText);
            slot.Value.Ancestors<Rich.RichValueData>().FirstOrDefault()?.Save();
            return true;
        }

        private bool TryRemoveExclusiveInCellImage(Cell cell) {
            if (!TryCreateInCellImageLookup(out InCellImageLookup? lookup)
                || !lookup!.TryResolveSlot(cell.ValueMetaIndex?.Value ?? 0U, out InCellImageSlot slot)
                || !IsExclusiveInCellImageSlot(lookup, slot)) {
                return false;
            }
            string? relationshipId = lookup.Relationships[slot.RelationshipIndex].Id?.Value;
            if (relationshipId == null) return false;
            OpenXmlPart imagePart;
            try {
                imagePart = lookup.RelationshipPart.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return false;
            }
            Rich.RichValueData? valuesRoot = slot.Value.Ancestors<Rich.RichValueData>().FirstOrDefault();
            FutureMetadata? futureRoot = lookup.FutureBlocks[slot.FutureIndex].Ancestors<FutureMetadata>().FirstOrDefault();
            ValueMetadata? valueMetadata = lookup.Metadata.GetFirstChild<ValueMetadata>();

            foreach (Rich.RichValue value in lookup.RichValues) {
                uint structureIndex = value.S?.Value ?? uint.MaxValue;
                if (structureIndex > int.MaxValue) continue;
                Rich.RichValueStructure? structure = lookup.Structures.ElementAtOrDefault((int)structureIndex);
                if (TryGetImageRelationshipIndex(value, structure, out int index, out _)
                    && index > slot.RelationshipIndex) {
                    SetImageRelationshipIndex(value, structure!, index - 1);
                }
            }
            lookup.RelationshipPart.DeletePart(imagePart);
            var relationships = new RichRel.RichValueRels(lookup.Relationships
                .Where((_, index) => index != slot.RelationshipIndex)
                .Select(item => item.CloneNode(true)));
            if (relationships.Elements<RichRel.RichValueRelRelationship>().Any()) {
                SaveExtendedRoot(lookup.RelationshipPart, relationships);
            } else {
                _excelDocument.WorkbookPartRoot.DeletePart(lookup.RelationshipPart);
            }

            foreach (FutureMetadataBlock block in lookup.FutureBlocks) {
                if (TryGetRichValueIndex(block, out uint index) && index > (uint)slot.ValueIndex) {
                    SetRichValueIndex(block, index - 1U);
                }
            }
            slot.Value.Remove();
            if (valuesRoot != null) {
                if (valuesRoot.Elements<Rich.RichValue>().Any()) {
                    valuesRoot.Count = (uint)valuesRoot.Elements<Rich.RichValue>().Count();
                    valuesRoot.Save();
                } else {
                    RdRichValuePart? valuePart = _excelDocument.WorkbookPartRoot.RdRichValueParts
                        .FirstOrDefault(part => ReferenceEquals(part.RichValueData, valuesRoot));
                    if (valuePart != null) _excelDocument.WorkbookPartRoot.DeletePart(valuePart);
                }
            }

            foreach (MetadataBlock block in lookup.MetadataBlocks) {
                foreach (MetadataRecord record in block.Elements<MetadataRecord>()) {
                    if (IsRichValueMetadataType(lookup.Metadata, record.TypeIndex?.Value ?? 0U)
                        && record.Val?.Value is uint index
                        && index > (uint)slot.FutureIndex) {
                        record.Val = index - 1U;
                    }
                }
            }
            lookup.FutureBlocks[slot.FutureIndex].Remove();
            if (futureRoot != null) {
                if (futureRoot.Elements<FutureMetadataBlock>().Any()) {
                    futureRoot.Count = (uint)futureRoot.Elements<FutureMetadataBlock>().Count();
                } else {
                    futureRoot.Remove();
                }
            }

            int removedMetadataIndex = checked((int)slot.MetadataIndex - 1);
            lookup.MetadataBlocks[removedMetadataIndex].Remove();
            if (valueMetadata != null) {
                if (valueMetadata.Elements<MetadataBlock>().Any()) {
                    valueMetadata.Count = (uint)valueMetadata.Elements<MetadataBlock>().Count();
                } else {
                    valueMetadata.Remove();
                }
            }
            foreach (WorksheetPart worksheetPart in _excelDocument.WorkbookPartRoot.WorksheetParts) {
                bool changed = false;
                foreach (Cell candidate in worksheetPart.Worksheet?.Descendants<Cell>() ?? Enumerable.Empty<Cell>()) {
                    if (candidate.ValueMetaIndex?.Value == slot.MetadataIndex) {
                        ClearCellValueMetadataAttribute(candidate);
                        changed = true;
                    } else if (candidate.ValueMetaIndex?.Value is uint index && index > slot.MetadataIndex) {
                        candidate.ValueMetaIndex = index - 1U;
                        changed = true;
                    }
                }
                if (changed) worksheetPart.Worksheet?.Save();
            }
            lookup.Metadata.Save();
            return true;
        }

        private static void SetRichValueIndex(FutureMetadataBlock block, uint valueIndex) =>
            RichValueErrorLookup.FindValueBlock(OfficeXmlNode.FromSdk(block)).SetUInt32Attribute("i", valueIndex);

        private static void SetImageRelationshipIndex(Rich.RichValue value, Rich.RichValueStructure structure, int relationshipIndex) {
            OfficeXmlNode field = FindRichImageField(OfficeXmlNode.FromSdk(value), OfficeXmlNode.FromSdk(structure), "_rvRel:LocalImageIdentifier");
            if (field.Exists) field.SetInnerText(relationshipIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        private bool IsExclusiveInCellImageSlot(InCellImageLookup lookup, InCellImageSlot slot) {
            int cellReferences = _excelDocument.WorkbookPartRoot.WorksheetParts
                .SelectMany(part => part.Worksheet?.Descendants<Cell>() ?? Enumerable.Empty<Cell>())
                .Count(cell => cell.ValueMetaIndex?.Value == slot.MetadataIndex);
            if (cellReferences != 1) return false;
            int futureReferences = lookup.MetadataBlocks.Count(block =>
                block.Elements<MetadataRecord>().Any(record =>
                    IsRichValueMetadataType(lookup.Metadata, record.TypeIndex?.Value ?? 0U)
                    && record.Val?.Value == (uint)slot.FutureIndex));
            if (futureReferences != 1) return false;
            int valueReferences = lookup.FutureBlocks.Count(block =>
                TryGetRichValueIndex(block, out uint index) && index == (uint)slot.ValueIndex);
            if (valueReferences != 1) return false;
            string? relationshipId = lookup.Relationships[slot.RelationshipIndex].Id?.Value;
            if (relationshipId == null
                || lookup.Relationships.Count(item => string.Equals(
                    item.Id?.Value,
                    relationshipId,
                    StringComparison.Ordinal)) != 1) {
                return false;
            }
            OpenXmlPart imagePart;
            try {
                imagePart = lookup.RelationshipPart.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return false;
            }
            OpenXmlPart[] imageParents = imagePart.GetParentParts().ToArray();
            if (imageParents.Length != 1 || !ReferenceEquals(imageParents[0], lookup.RelationshipPart)) {
                return false;
            }
            int relationshipReferences = 0;
            foreach (Rich.RichValue value in lookup.RichValues) {
                uint structureIndex = value.S?.Value ?? uint.MaxValue;
                if (structureIndex > int.MaxValue) continue;
                Rich.RichValueStructure? structure = lookup.Structures.ElementAtOrDefault((int)structureIndex);
                if (TryGetImageRelationshipIndex(value, structure, out int index, out _)
                    && index == slot.RelationshipIndex) {
                    relationshipReferences++;
                }
            }
            return relationshipReferences == 1;
        }

        private static bool TryGetRichValueIndex(FutureMetadataBlock? futureBlock, out uint valueIndex) {
            valueIndex = RichValueErrorLookup.FindValueBlock(OfficeXmlNode.FromSdk(futureBlock)).UInt32Attribute("i") ?? uint.MaxValue;
            return valueIndex != uint.MaxValue;
        }

        private static bool TryGetImageRelationshipIndex(Rich.RichValue? value, Rich.RichValueStructure? structure,
            out int relationshipIndex, out string altText) =>
            TryGetImageRelationshipIndex(OfficeXmlNode.FromSdk(value), OfficeXmlNode.FromSdk(structure), out relationshipIndex, out altText);

        private static bool TryGetImageRelationshipIndex(OfficeXmlNode value, OfficeXmlNode structure,
            out int relationshipIndex, out string altText) {
            relationshipIndex = -1;
            altText = string.Empty;
            if (!string.Equals(structure.Attribute("t"), "_localImage", StringComparison.OrdinalIgnoreCase)) return false;
            OfficeXmlNode identifier = FindRichImageField(value, structure, "_rvRel:LocalImageIdentifier");
            if (!uint.TryParse(identifier.InnerText, out uint index) || index > int.MaxValue) return false;
            relationshipIndex = (int)index;
            altText = FindRichImageField(value, structure, "Text").InnerText;
            return true;
        }

        private static void SetRichValueAltText(Rich.RichValue value, Rich.RichValueStructure structure, string altText) {
            OfficeXmlNode text = FindRichImageField(OfficeXmlNode.FromSdk(value), OfficeXmlNode.FromSdk(structure), "Text");
            if (text.Exists) text.SetInnerText(altText);
        }

        private static OfficeXmlNode FindRichImageField(OfficeXmlNode value, OfficeXmlNode structure, string fieldName) {
            const string rich = RichValueErrorLookup.RichData;
            OfficeXmlNode field = value.FirstChild;
            foreach (OfficeXmlNode key in structure.Children()) {
                if (!key.Is(rich, "k")) continue;
                while (field.Exists && !field.Is(rich, "v")) field = field.NextSibling;
                if (field.IsNone || string.Equals(key.Attribute("n"), fieldName, StringComparison.OrdinalIgnoreCase)) return field;
                field = field.NextSibling;
            }
            return default;
        }
    }
}
