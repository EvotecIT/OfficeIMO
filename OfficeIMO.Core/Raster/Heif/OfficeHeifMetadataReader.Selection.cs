// Adapted from the Evotec ImagePlayground HEIF metadata implementation.
// Copyright (c) 2022 Evotec. MIT license; see Licenses/ImagePlayground-LICENSE.txt.
using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeHeifMetadataReader {
    private sealed partial class Parser {
        private bool TryFindMetadataItemId(byte[] data, Box metaBox, Box itemInfoBox,
            bool exif, bool requireSupportedPayload, out uint itemId) {
            itemId = 0;
            var items = new List<HeifItemInfoBuilder>();
            if (!TryReadItemInfos(data, itemInfoBox, items)) {
                return false;
            }
            ReadMetadataAssociations(data, metaBox, out uint? primary, out List<OfficeHeifItemReference> references);
            if (!TrySelectMetadataItem(items, references, primary, exif, out HeifItemInfoBuilder? selected)) {
                return false;
            }
            if (requireSupportedPayload && (selected!.ItemProtectionIndex != 0 ||
                (!exif && !string.IsNullOrEmpty(selected.ContentEncoding)))) {
                return false;
            }
            itemId = selected!.ItemId;
            return true;
        }

        private bool HasMetadataItem(byte[] data, Box itemInfoBox, bool exif) {
            var items = new List<HeifItemInfoBuilder>();
            if (!TryReadItemInfos(data, itemInfoBox, items)) {
                return false;
            }
            foreach (HeifItemInfoBuilder item in items) {
                CheckWork();
                if (IsMetadataFamily(item, exif)) {
                    return true;
                }
            }
            return false;
        }

        private void ReadMetadataAssociations(byte[] data, Box metaBox, out uint? primary,
            out List<OfficeHeifItemReference> references) {
            primary = null;
            references = new List<OfficeHeifItemReference>();
            bool foundReferences = false;
            foreach (Box child in EnumerateBoxes(data, metaBox.DataOffset + 4, metaBox.EndOffset)) {
                CheckWork();
                if (child.Type == "pitm") {
                    if (primary.HasValue || !TryReadPrimaryItemId(data, child, out uint value)) {
                        throw new FormatException("Invalid HEIF primary item declaration.");
                    }
                    primary = value;
                } else if (child.Type == "iref") {
                    if (foundReferences) {
                        throw new FormatException("Duplicate HEIF reference collection.");
                    }
                    references = ReadItemReferences(data, child);
                    foundReferences = true;
                }
            }
        }

        private bool TrySelectMetadataItem(IReadOnlyList<HeifItemInfoBuilder> items,
            IReadOnlyList<OfficeHeifItemReference> references, uint? primary, bool exif,
            out HeifItemInfoBuilder? selected) {
            selected = null;
            var candidates = new Dictionary<uint, HeifItemInfoBuilder>();
            bool primaryDeclared = false;
            foreach (HeifItemInfoBuilder item in items) {
                CheckWork();
                primaryDeclared |= primary.HasValue && item.ItemId == primary.Value;
                if (IsMetadataFamily(item, exif)) {
                    candidates.Add(item.ItemId, item);
                }
            }
            if (candidates.Count == 0 || (primary.HasValue && !primaryDeclared)) {
                return false;
            }
            var associated = new HashSet<uint>();
            bool hasExplicitAssociation = false;
            foreach (OfficeHeifItemReference reference in references) {
                CheckWork();
                if (reference.ReferenceType != "cdsc" || !candidates.ContainsKey(reference.FromItemId)) {
                    continue;
                }
                foreach (uint target in reference.ToItemIds) {
                    CheckWork();
                    hasExplicitAssociation = true;
                    if (primary.HasValue && target == primary.Value) {
                        associated.Add(reference.FromItemId);
                    }
                }
            }
            if (associated.Count == 1) {
                foreach (uint id in associated) {
                    selected = candidates[id];
                }
                return true;
            }
            if (associated.Count > 1 || candidates.Count != 1 ||
                (primary.HasValue && hasExplicitAssociation)) {
                return false;
            }
            // Older unique-item containers do not necessarily declare pitm/cdsc.
            foreach (HeifItemInfoBuilder item in candidates.Values) {
                selected = item;
            }
            return true;
        }

        private bool IsMetadataFamily(HeifItemInfoBuilder item, bool exif) =>
            exif ? item.ItemType == ExifItemType : IsXmpMimeType(item.MimeType);
    }
}
