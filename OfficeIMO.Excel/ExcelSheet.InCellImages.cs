using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing.Internal;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;
using RichRel = DocumentFormat.OpenXml.Office.Y2022.RichValueRel;

namespace OfficeIMO.Excel {
    /// <summary>Native image stored as an Excel rich value in a worksheet cell.</summary>
    public sealed class ExcelInCellImage {
        internal ExcelInCellImage(string cellReference, string contentType, string altText, byte[] bytes) {
            CellReference = cellReference;
            ContentType = contentType;
            AltText = altText;
            Bytes = bytes;
        }
        /// <summary>Owning A1 cell.</summary>
        public string CellReference { get; }
        /// <summary>Image MIME type.</summary>
        public string ContentType { get; }
        /// <summary>Accessible alternative text.</summary>
        public string AltText { get; }
        /// <summary>Stored image payload.</summary>
        public byte[] Bytes { get; }
    }

    public partial class ExcelSheet {
        private const string RichValueMetadataName = "XLRICHVALUE";
        private const string RichValueMetadataExtensionUri = "{3E2802C4-A4D2-4D8B-9148-E3BE6C30E623}";
        private const string RichValueRelRelationshipType = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
        private const string RichValueRelContentType = "application/vnd.ms-excel.richvaluerel+xml";
        private const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        private const long MaximumRichValueMetadataBytes = 16L * 1024L * 1024L;
        private const long MaximumRichValueRelationshipBytes = 16L * 1024L * 1024L;

        internal static void ValidateInCellImageMetadataStream(
            Stream source,
            long maximumBytes,
            string description) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            if (string.IsNullOrWhiteSpace(description)) throw new ArgumentException("Metadata description cannot be empty.", nameof(description));

            if (source.CanSeek) {
                long position = source.Position;
                try {
                    if (source.Length > maximumBytes) {
                        throw new InvalidDataException(
                            $"{description} exceeds the supported {maximumBytes}-byte limit.");
                    }
                } finally {
                    source.Position = position;
                }
                return;
            }

            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) return;
                total += read;
                if (total > maximumBytes) {
                    throw new InvalidDataException(
                        $"{description} exceeds the supported {maximumBytes}-byte limit.");
                }
            }
        }

        private static void ValidateInCellImageMetadataPart(OpenXmlPart part, string description) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            ValidateInCellImageMetadataStream(stream, MaximumRichValueMetadataBytes, description);
        }

        /// <summary>Adds or replaces a native in-cell image without creating a drawing anchor.</summary>
        public ExcelInCellImage SetInCellImage(
            int row,
            int column,
            byte[] imageBytes,
            string contentType = "image/png",
            string? altText = null,
            long maximumImageBytes = 32_000_000) {
            if (row < 1 || row > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1 || column > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
            if (imageBytes == null) throw new ArgumentNullException(nameof(imageBytes));
            if (imageBytes.LongLength == 0 || imageBytes.LongLength > maximumImageBytes) throw new ArgumentOutOfRangeException(nameof(imageBytes));
            if (!IsSupportedImageContentType(contentType)) throw new NotSupportedException($"Image content type '{contentType}' is not supported.");
            string text = altText ?? string.Empty;
            ValidateInCellImageAltText(text);
            WriteLock(() => {
                using var imageStream = new MemoryStream(imageBytes, writable: false);
                SetInCellImageCore(row, column, imageStream, contentType, text);
                CompleteCellValueMutation(row, column);
                WorksheetRoot.Save();
            });
            return new ExcelInCellImage(A1.CellReference(row, column), contentType, text, (byte[])imageBytes.Clone());
        }

        private static void ValidateInCellImageAltText(string altText) {
            if (altText.Length > 32_767) {
                throw new ArgumentException("In-cell image alternative text exceeds Excel's 32,767-character limit.", nameof(altText));
            }
            try {
                XmlConvert.VerifyXmlChars(altText);
            } catch (XmlException ex) {
                throw new ArgumentException("In-cell image alternative text must contain valid Excel XML text.", nameof(altText), ex);
            }
        }

        private void SetInCellImageCore(int row, int column, Stream imageStream, string contentType, string altText) {
            _ = SetInCellImageCore(row, column, imageStream, contentType, altText, reusableImagePart: null);
        }

        private OpenXmlPart SetInCellImageCore(
            int row,
            int column,
            Stream? imageStream,
            string contentType,
            string altText,
            OpenXmlPart? reusableImagePart) {
            Cell cell = GetCell(row, column);
            if (reusableImagePart == null
                && cell.ValueMetaIndex != null
                && TryReplaceExclusiveInCellImage(
                    cell,
                    imageStream!,
                    contentType,
                    altText,
                    out OpenXmlPart? replacedImagePart)) {
                SetInCellImageCellValue(cell);
                return replacedImagePart!;
            }

            WorkbookPart workbookPart = _excelDocument.WorkbookPartRoot;
            CellMetadataPart metadataPart = workbookPart.CellMetadataPart ?? workbookPart.AddNewPart<CellMetadataPart>();
            Metadata metadata = metadataPart.Metadata ??= new Metadata();
            uint typeIndex = EnsureRichValueMetadataType(metadata);
            FutureMetadata future = EnsureRichValueFutureMetadata(metadata);
            ValueMetadata valueMetadata = metadata.GetFirstChild<ValueMetadata>() ?? metadata.AppendChild(new ValueMetadata());

            RdRichValuePart valuePart = workbookPart.RdRichValueParts.FirstOrDefault()
                ?? workbookPart.AddNewPart<RdRichValuePart>();
            Rich.RichValueData values = valuePart.RichValueData ??= new Rich.RichValueData();
            RdRichValueStructurePart structurePart = workbookPart.GetPartsOfType<RdRichValueStructurePart>().FirstOrDefault()
                ?? workbookPart.AddNewPart<RdRichValueStructurePart>();
            Rich.RichValueStructures structures = structurePart.RichValueStructures ??= new Rich.RichValueStructures();
            uint structureIndex = EnsureLocalImageStructure(structures);

            ExtendedPart relationshipPart = workbookPart.Parts.Select(pair => pair.OpenXmlPart).OfType<ExtendedPart>()
                .FirstOrDefault(part => string.Equals(part.RelationshipType, RichValueRelRelationshipType, StringComparison.Ordinal))
                ?? workbookPart.AddExtendedPart(RichValueRelRelationshipType, RichValueRelContentType, "xml");
            RichRel.RichValueRels relationships = LoadRichValueRelationships(relationshipPart);
            OpenXmlPart resolvedImagePart;
            string imageRelationshipId;
            uint relationshipIndex;
            if (reusableImagePart != null) {
                if (!relationshipPart.Parts.Any(pair => ReferenceEquals(pair.OpenXmlPart, reusableImagePart))) {
                    throw new InvalidOperationException("The reusable in-cell image asset is not owned by the target rich-value relationship part.");
                }
                resolvedImagePart = reusableImagePart;
                imageRelationshipId = relationshipPart.GetIdOfPart(reusableImagePart);
                List<RichRel.RichValueRelRelationship> existingRelationships = relationships
                    .Elements<RichRel.RichValueRelRelationship>()
                    .ToList();
                int existingIndex = existingRelationships.FindIndex(item =>
                    string.Equals(item.Id?.Value, imageRelationshipId, StringComparison.Ordinal));
                if (existingIndex >= 0) {
                    relationshipIndex = (uint)existingIndex;
                } else {
                    relationshipIndex = (uint)existingRelationships.Count;
                    relationships.Append(new RichRel.RichValueRelRelationship { Id = imageRelationshipId });
                    SaveExtendedRoot(relationshipPart, relationships);
                }
            } else {
                if (imageStream == null) throw new ArgumentNullException(nameof(imageStream));
                ExtendedPart imagePart = relationshipPart.AddExtendedPart(ImageRelationshipType, contentType, GetImageExtension(contentType));
                imagePart.FeedData(imageStream);
                resolvedImagePart = imagePart;
                imageRelationshipId = relationshipPart.GetIdOfPart(imagePart);
                relationshipIndex = (uint)relationships.Elements<RichRel.RichValueRelRelationship>().Count();
                relationships.Append(new RichRel.RichValueRelRelationship { Id = imageRelationshipId });
                SaveExtendedRoot(relationshipPart, relationships);
            }

            uint valueIndex = (uint)values.Elements<Rich.RichValue>().Count();
            Rich.RichValue value = CreateLocalImageValue(structureIndex, relationshipIndex, altText, structures);
            values.Append(value);
            values.Count = (uint)values.Elements<Rich.RichValue>().Count();
            values.Save();
            structures.Count = (uint)structures.Elements<Rich.RichValueStructure>().Count();
            structures.Save();

            uint futureIndex = (uint)future.Elements<FutureMetadataBlock>().Count();
            var extension = new Extension { Uri = RichValueMetadataExtensionUri };
            extension.Append(new Rich.RichValueBlock { I = valueIndex });
            future.Append(new FutureMetadataBlock(new ExtensionList(extension)));
            future.Count = (uint)future.Elements<FutureMetadataBlock>().Count();

            var metadataBlock = new MetadataBlock(new MetadataRecord { TypeIndex = typeIndex, Val = futureIndex });
            valueMetadata.Append(metadataBlock);
            valueMetadata.Count = (uint)valueMetadata.Elements<MetadataBlock>().Count();
            SetInCellImageCellValue(cell);
            cell.ValueMetaIndex = valueMetadata.Count;
            metadata.Save();
            return resolvedImagePart;
        }

        internal void PreflightInCellImages(ExcelDocument.InCellImageCopyContext context) {
            if (!TryCreateInCellImageLookup(out InCellImageLookup? lookup)) return;
            foreach (Cell sourceCell in WorksheetRoot.Descendants<Cell>()) {
                if (!lookup!.TryResolve(sourceCell, out OpenXmlPart? imagePart, out _)
                    || imagePart == null) {
                    continue;
                }

                if (!context.TryGetSourcePayload(imagePart, out byte[] payload)) {
                    using Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                    long maximumReadableBytes = context.GetMaximumReadableBytes();
                    payload = ReadInCellImagePayload(
                        imageStream,
                        maximumReadableBytes,
                        maximumReadableBytes);
                    context.AddSourcePayload(imagePart, payload);
                }
                context.Consume(payload.LongLength);
            }
        }

        internal void CopyInCellImagesTo(
            ExcelSheet targetSheet,
            ExcelDocument.InCellImageCopyContext context) {
            if (!TryCreateInCellImageLookup(out InCellImageLookup? lookup)) return;
            bool copied = false;
            foreach (Cell sourceCell in WorksheetRoot.Descendants<Cell>()) {
                if (!lookup!.TryResolve(sourceCell, out OpenXmlPart? imagePart, out string altText)
                    || imagePart == null
                    || sourceCell.CellReference?.Value is not string cellReference) {
                    continue;
                }

                (int row, int column) = A1.ParseCellRef(cellReference);
                if (context.TryGetCopiedAsset(imagePart, out OpenXmlPart copiedAsset)) {
                    _ = targetSheet.SetInCellImageCore(
                        row,
                        column,
                        imageStream: null,
                        imagePart.ContentType,
                        altText,
                        copiedAsset);
                } else {
                    if (!context.TryGetSourcePayload(imagePart, out byte[] payload)) {
                        throw new InvalidOperationException("In-cell images must be preflighted before package copy.");
                    }
                    using var payloadStream = new MemoryStream(payload, writable: false);
                    OpenXmlPart targetImagePart = targetSheet.SetInCellImageCore(
                        row,
                        column,
                        payloadStream,
                        imagePart.ContentType,
                        altText,
                        reusableImagePart: null);
                    context.AddCopiedAsset(imagePart, targetImagePart);
                }
                copied = true;
            }

            if (copied) {
                targetSheet.WorksheetRoot.Save();
            }
        }

        internal IReadOnlyList<string> GetInCellImageCellReferences() {
            if (!TryCreateInCellImageLookup(out InCellImageLookup? lookup)) return Array.Empty<string>();
            return WorksheetRoot.Descendants<Cell>()
                .Where(cell => lookup!.TryResolve(cell, out _, out _))
                .Select(cell => cell.CellReference?.Value ?? string.Empty)
                .Where(reference => reference.Length > 0)
                .ToArray();
        }

        /// <summary>Reads native in-cell images with a deterministic aggregate payload budget.</summary>
        public IReadOnlyList<ExcelInCellImage> GetInCellImages(long maximumTotalImageBytes = 64_000_000) {
            if (maximumTotalImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumTotalImageBytes));
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var result = new List<ExcelInCellImage>();
                if (!TryCreateInCellImageLookup(out InCellImageLookup? lookup)) {
                    return new ReadOnlyCollection<ExcelInCellImage>(result);
                }
                var payloads = new Dictionary<OpenXmlPart, byte[]>();
                long total = 0;
                foreach (Cell cell in WorksheetRoot.Descendants<Cell>()) {
                    if (!lookup!.TryResolve(cell, out OpenXmlPart? imagePart, out string altText) || imagePart == null) continue;
                    bool sharedPayload = payloads.TryGetValue(imagePart, out byte[]? payload);
                    if (!sharedPayload) {
                        using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                        payload = ReadInCellImagePayload(
                            source,
                            maximumTotalImageBytes - total,
                            maximumTotalImageBytes);
                        payloads.Add(imagePart, payload);
                    }
                    byte[] resolvedPayload = payload!;
                    total = checked(total + resolvedPayload.LongLength);
                    if (total > maximumTotalImageBytes) {
                        throw new InvalidOperationException($"In-cell image payloads exceed maximumTotalImageBytes ({maximumTotalImageBytes}).");
                    }
                    result.Add(new ExcelInCellImage(
                        cell.CellReference?.Value ?? string.Empty,
                        imagePart.ContentType,
                        altText,
                        sharedPayload ? (byte[])resolvedPayload.Clone() : resolvedPayload));
                }
                return new ReadOnlyCollection<ExcelInCellImage>(result);
            });
        }

        /// <summary>Reads one image without allowing a non-seekable package stream to exceed the remaining aggregate budget.</summary>
        internal static byte[] ReadInCellImagePayload(
            Stream source,
            long remainingAggregateBytes,
            long maximumTotalImageBytes) {
            if (ExcelImageExportLimits.TryReadSourceImageBytes(source, remainingAggregateBytes, out byte[] payload)) {
                return payload;
            }
            throw new InvalidOperationException(
                $"In-cell image payloads exceed maximumTotalImageBytes ({maximumTotalImageBytes}).");
        }

        /// <summary>Removes the in-cell image value from one cell. Shared rich-data assets remain available to copied cells.</summary>
        public bool RemoveInCellImage(int row, int column) {
            bool removed = false;
            WriteLock(() => {
                Cell? cell = TryGetExistingCell(row, column);
                if (cell == null || !TryResolveInCellImage(cell, out _, out _)) return;
                if (!TryRemoveExclusiveInCellImage(cell)) ClearCellValueMetadataAttribute(cell);
                cell.CellValue = null;
                cell.DataType = null;
                cell.InlineString = null;
                CompleteCellValueMutation(row, column);
                WorksheetRoot.Save();
                removed = true;
            });
            return removed;
        }

        private static bool HasCellValueMetadata(Cell cell) => cell.ValueMetaIndex != null;

        private void ClearCellValueMetadata(Cell cell) {
            if (cell.ValueMetaIndex == null) return;
            if (TryRemoveExclusiveInCellImage(cell)) return;
            ClearCellValueMetadataAttribute(cell);
        }

        private static void ClearCellValueMetadataAttribute(Cell cell) {
            cell.ValueMetaIndex = null;
            cell.RemoveAttribute("vm", string.Empty);
        }

        private void RemoveCellWithValueMetadataCleanup(Cell cell) {
            ClearCellValueMetadata(cell);
            cell.Remove();
        }

        private bool TryResolveInCellImage(Cell cell, out OpenXmlPart? imagePart, out string altText) {
            if (!TryCreateInCellImageLookup(out InCellImageLookup? lookup)) {
                imagePart = null;
                altText = string.Empty;
                return false;
            }
            return lookup!.TryResolve(cell, out imagePart, out altText);
        }

        private static void SetInCellImageCellValue(Cell cell) {
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
            cell.CellValue = new CellValue("#VALUE!");
            cell.CellFormula = null;
            cell.InlineString = null;
        }

        private static uint EnsureRichValueMetadataType(Metadata metadata) {
            MetadataTypes types = metadata.MetadataTypes ??= new MetadataTypes();
            List<MetadataType> existing = types.Elements<MetadataType>().ToList();
            int index = existing.FindIndex(type => string.Equals(type.Name?.Value, RichValueMetadataName, StringComparison.OrdinalIgnoreCase));
            if (index < 0) {
                types.Append(new MetadataType {
                    Name = RichValueMetadataName,
                    MinSupportedVersion = 120000U,
                    Copy = true,
                    PasteAll = true,
                    PasteValues = true,
                    Merge = true,
                    SplitFirst = true,
                    RowColumnShift = true,
                    ClearAll = true,
                    ClearContents = true
                });
                index = existing.Count;
            }
            types.Count = (uint)types.Elements<MetadataType>().Count();
            return (uint)index + 1U;
        }

        private static FutureMetadata EnsureRichValueFutureMetadata(Metadata metadata) {
            FutureMetadata? future = metadata.Elements<FutureMetadata>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, RichValueMetadataName, StringComparison.OrdinalIgnoreCase));
            if (future != null) return future;
            future = new FutureMetadata { Name = RichValueMetadataName, Count = 0U };
            OpenXmlElement? metadataBlocks = metadata.ChildElements
                .FirstOrDefault(element => element is CellMetadata || element is ValueMetadata);
            if (metadataBlocks == null) metadata.Append(future); else metadata.InsertBefore(future, metadataBlocks);
            return future;
        }

        private static uint EnsureLocalImageStructure(Rich.RichValueStructures structures) {
            List<Rich.RichValueStructure> existing = structures.Elements<Rich.RichValueStructure>().ToList();
            int index = existing.FindIndex(structure => string.Equals(structure.T?.Value, "_localImage", StringComparison.OrdinalIgnoreCase));
            if (index >= 0) return (uint)index;
            structures.Append(new Rich.RichValueStructure(
                new Rich.Key { N = "_rvRel:LocalImageIdentifier", T = Rich.RichValueValueType.I },
                new Rich.Key { N = "CalcOrigin", T = Rich.RichValueValueType.I },
                new Rich.Key { N = "Text", T = Rich.RichValueValueType.S }) { T = "_localImage" });
            return (uint)existing.Count;
        }

        private static Rich.RichValue CreateLocalImageValue(uint structureIndex, uint relationshipIndex, string altText, Rich.RichValueStructures structures) {
            Rich.RichValueStructure structure = structures.Elements<Rich.RichValueStructure>().ElementAt((int)structureIndex);
            var value = new Rich.RichValue { S = structureIndex };
            foreach (Rich.Key key in structure.Elements<Rich.Key>()) {
                string name = key.N?.Value ?? string.Empty;
                value.Append(new Rich.Value(
                    string.Equals(name, "_rvRel:LocalImageIdentifier", StringComparison.OrdinalIgnoreCase) ? relationshipIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : string.Equals(name, "CalcOrigin", StringComparison.OrdinalIgnoreCase) ? "5"
                    : string.Equals(name, "Text", StringComparison.OrdinalIgnoreCase) ? altText
                    : string.Empty));
            }
            return value;
        }

        private static bool IsRichValueMetadataType(Metadata metadata, uint oneBasedIndex) {
            if (oneBasedIndex == 0U) return false;
            MetadataType? type = metadata.MetadataTypes?.Elements<MetadataType>().ElementAtOrDefault((int)oneBasedIndex - 1);
            return string.Equals(type?.Name?.Value, RichValueMetadataName, StringComparison.OrdinalIgnoreCase);
        }

        private static RichRel.RichValueRels LoadRichValueRelationships(ExtendedPart part) {
            using Stream stream = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
            if (stream.Length == 0) return new RichRel.RichValueRels();
            byte[] bytes;
            try {
                bytes = OfficeStreamReader.ReadAllBytes(stream, MaximumRichValueRelationshipBytes);
            } catch (InvalidDataException exception) {
                throw new InvalidDataException(
                    $"Rich-value relationship metadata exceeds the supported {MaximumRichValueRelationshipBytes}-byte limit.",
                    exception);
            }
            using var bounded = new MemoryStream(bytes, writable: false);
            using var reader = new StreamReader(bounded, Encoding.UTF8, true, 1024, leaveOpen: false);
            return new RichRel.RichValueRels(reader.ReadToEnd());
        }

        private static void SaveExtendedRoot(ExtendedPart part, OpenXmlElement root) {
            using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024, leaveOpen: false);
            writer.Write(root.OuterXml);
        }

        private static string GetImageExtension(string contentType) => contentType.ToLowerInvariant() switch {
            "image/jpeg" => "jpeg",
            "image/jpg" => "jpg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            "image/svg+xml" => "svg",
            _ => "png"
        };
    }
}
