using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
            WriteLock(() => {
                using var imageStream = new MemoryStream(imageBytes, writable: false);
                SetInCellImageCore(row, column, imageStream, contentType, text);
                WorksheetRoot.Save();
            });
            return new ExcelInCellImage(A1.CellReference(row, column), contentType, text, (byte[])imageBytes.Clone());
        }

        private void SetInCellImageCore(int row, int column, Stream imageStream, string contentType, string altText) {
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
            ExtendedPart imagePart = relationshipPart.AddExtendedPart(ImageRelationshipType, contentType, GetImageExtension(contentType));
            imagePart.FeedData(imageStream);
            string imageRelationshipId = relationshipPart.GetIdOfPart(imagePart);
            uint relationshipIndex = (uint)relationships.Elements<RichRel.RichValueRelRelationship>().Count();
            relationships.Append(new RichRel.RichValueRelRelationship { Id = imageRelationshipId });
            SaveExtendedRoot(relationshipPart, relationships);

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
            Cell cell = GetCell(row, column);
            cell.ValueMetaIndex = valueMetadata.Count;
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
            cell.CellValue = new CellValue("#VALUE!");
            cell.CellFormula = null;
            cell.InlineString = null;
            metadata.Save();
        }

        internal void CopyInCellImagesTo(ExcelSheet targetSheet) {
            bool copied = false;
            foreach (Cell sourceCell in WorksheetRoot.Descendants<Cell>()) {
                if (!TryResolveInCellImage(sourceCell, out OpenXmlPart? imagePart, out string altText)
                    || imagePart == null
                    || sourceCell.CellReference?.Value is not string cellReference) {
                    continue;
                }

                (int row, int column) = A1.ParseCellRef(cellReference);
                using Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                targetSheet.SetInCellImageCore(row, column, imageStream, imagePart.ContentType, altText);
                copied = true;
            }

            if (copied) {
                targetSheet.WorksheetRoot.Save();
            }
        }

        /// <summary>Reads native in-cell images with a deterministic aggregate payload budget.</summary>
        public IReadOnlyList<ExcelInCellImage> GetInCellImages(long maximumTotalImageBytes = 64_000_000) {
            if (maximumTotalImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumTotalImageBytes));
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var result = new List<ExcelInCellImage>();
                long total = 0;
                foreach (Cell cell in WorksheetRoot.Descendants<Cell>()) {
                    if (!TryResolveInCellImage(cell, out OpenXmlPart? imagePart, out string altText) || imagePart == null) continue;
                    using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                    if (source.CanSeek && checked(total + source.Length) > maximumTotalImageBytes) {
                        throw new InvalidOperationException($"In-cell image payloads exceed maximumTotalImageBytes ({maximumTotalImageBytes}).");
                    }
                    using var buffer = new MemoryStream();
                    source.CopyTo(buffer);
                    total = checked(total + buffer.Length);
                    if (total > maximumTotalImageBytes) throw new InvalidOperationException($"In-cell image payloads exceed maximumTotalImageBytes ({maximumTotalImageBytes}).");
                    result.Add(new ExcelInCellImage(cell.CellReference?.Value ?? string.Empty, imagePart.ContentType, altText, buffer.ToArray()));
                }
                return new ReadOnlyCollection<ExcelInCellImage>(result);
            });
        }

        /// <summary>Removes the in-cell image value from one cell. Shared rich-data assets remain available to copied cells.</summary>
        public bool RemoveInCellImage(int row, int column) {
            bool removed = false;
            WriteLock(() => {
                Cell? cell = TryGetExistingCell(row, column);
                if (cell == null || !TryResolveInCellImage(cell, out _, out _)) return;
                cell.ValueMetaIndex = null;
                cell.RemoveAttribute("vm", string.Empty);
                cell.CellValue = null;
                cell.DataType = null;
                cell.InlineString = null;
                WorksheetRoot.Save();
                removed = true;
            });
            return removed;
        }

        private static bool HasCellValueMetadata(Cell cell) => cell.ValueMetaIndex != null;

        private static void ClearCellValueMetadata(Cell cell) {
            if (cell.ValueMetaIndex == null) return;
            cell.ValueMetaIndex = null;
            cell.RemoveAttribute("vm", string.Empty);
        }

        private bool TryResolveInCellImage(Cell cell, out OpenXmlPart? imagePart, out string altText) {
            imagePart = null;
            altText = string.Empty;
            uint metadataIndex = cell.ValueMetaIndex?.Value ?? 0U;
            Metadata? metadata = _excelDocument.WorkbookPartRoot.CellMetadataPart?.Metadata;
            ValueMetadata? valueMetadata = metadata?.GetFirstChild<ValueMetadata>();
            if (metadataIndex == 0U || metadata == null || valueMetadata == null) return false;
            MetadataBlock? block = valueMetadata.Elements<MetadataBlock>().ElementAtOrDefault((int)metadataIndex - 1);
            MetadataRecord? record = block?.Elements<MetadataRecord>().FirstOrDefault();
            if (record == null || !IsRichValueMetadataType(metadata, record.TypeIndex?.Value ?? 0U)) return false;
            FutureMetadata? future = metadata.Elements<FutureMetadata>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, RichValueMetadataName, StringComparison.OrdinalIgnoreCase));
            FutureMetadataBlock? futureBlock = future?.Elements<FutureMetadataBlock>().ElementAtOrDefault((int)(record.Val?.Value ?? uint.MaxValue));
            OpenXmlElement? valueBlock = futureBlock?.Descendants<Rich.RichValueBlock>().FirstOrDefault()
                ?? futureBlock?.Descendants().FirstOrDefault(element =>
                    string.Equals(element.LocalName, "rvb", StringComparison.Ordinal));
            uint valueIndex = uint.MaxValue;
            if (valueBlock is Rich.RichValueBlock typedBlock) {
                valueIndex = typedBlock.I?.Value ?? uint.MaxValue;
            } else if (valueBlock != null) {
                uint.TryParse(valueBlock.GetAttribute("i", string.Empty).Value,
                    System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out valueIndex);
            }
            WorkbookPart workbookPart = _excelDocument.WorkbookPartRoot;
            RdRichValuePart? valuePart = workbookPart.RdRichValueParts.FirstOrDefault();
            Rich.RichValue? value = valuePart?.RichValueData?.Elements<Rich.RichValue>().ElementAtOrDefault((int)valueIndex);
            RdRichValueStructurePart? structurePart = workbookPart.GetPartsOfType<RdRichValueStructurePart>().FirstOrDefault();
            Rich.RichValueStructure? structure = structurePart?.RichValueStructures?.Elements<Rich.RichValueStructure>().ElementAtOrDefault((int)(value?.S?.Value ?? uint.MaxValue));
            if (!string.Equals(structure?.T?.Value, "_localImage", StringComparison.OrdinalIgnoreCase)) return false;
            List<Rich.Key> keys = structure!.Elements<Rich.Key>().ToList();
            List<Rich.Value> values = value!.Elements<Rich.Value>().ToList();
            int identifierIndex = keys.FindIndex(key => string.Equals(key.N?.Value, "_rvRel:LocalImageIdentifier", StringComparison.OrdinalIgnoreCase));
            int textIndex = keys.FindIndex(key => string.Equals(key.N?.Value, "Text", StringComparison.OrdinalIgnoreCase));
            if (identifierIndex < 0 || identifierIndex >= values.Count || !uint.TryParse(values[identifierIndex].Text, out uint relationIndex)) return false;
            if (textIndex >= 0 && textIndex < values.Count) altText = values[textIndex].Text;
            ExtendedPart? relationshipPart = workbookPart.Parts.Select(pair => pair.OpenXmlPart).OfType<ExtendedPart>()
                .FirstOrDefault(part => string.Equals(part.RelationshipType, RichValueRelRelationshipType, StringComparison.Ordinal));
            if (relationshipPart == null) return false;
            RichRel.RichValueRelRelationship? relationship = LoadRichValueRelationships(relationshipPart)
                .Elements<RichRel.RichValueRelRelationship>().ElementAtOrDefault((int)relationIndex);
            if (relationship?.Id?.Value is not string relationshipId) return false;
            try { imagePart = relationshipPart.GetPartById(relationshipId); } catch (ArgumentOutOfRangeException) { return false; }
            return imagePart != null && imagePart.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
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
            using var reader = new StreamReader(stream, Encoding.UTF8, true, 1024, leaveOpen: false);
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
