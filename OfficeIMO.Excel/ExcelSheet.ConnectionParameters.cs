using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string WorkbookConnectionRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
        private const string WorkbookConnectionContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";

        private sealed class WorkbookConnectionRoot {
            internal WorkbookConnectionRoot(OpenXmlPart part, Connections connections) {
                Part = part;
                Connections = connections;
            }

            internal OpenXmlPart Part { get; }
            internal Connections Connections { get; }

            internal void Save() {
                if (Part is ConnectionsPart nativePart) {
                    nativePart.Connections?.Save();
                    return;
                }

                string xml = Connections.OuterXml;
                if (xml.Length > ExcelDocument.MaximumWorkbookConnectionMetadataCharacters) {
                    throw new InvalidDataException(
                        $"Workbook connection metadata exceeds {ExcelDocument.MaximumWorkbookConnectionMetadataCharacters} characters.");
                }
                using var payload = new MemoryStream(Encoding.UTF8.GetBytes(xml), writable: false);
                Part.FeedData(payload);
            }
        }

        private IReadOnlyList<WorkbookConnectionRoot> LoadWorkbookConnectionRoots(
            MutationPlanScanBudget? budget = null) {
            var result = new List<WorkbookConnectionRoot>();
            foreach (IdPartPair pair in WorkbookPartRoot.Parts) {
                OpenXmlPart part = pair.OpenXmlPart;
                bool isNative = part is ConnectionsPart;
                if (!isNative
                    && (part is not ExtendedPart
                    || (!string.Equals(part.RelationshipType, WorkbookConnectionRelationshipType, StringComparison.Ordinal)
                        && !string.Equals(part.ContentType, WorkbookConnectionContentType, StringComparison.OrdinalIgnoreCase)))) {
                    continue;
                }

                if (part is ConnectionsPart loadedNative && loadedNative.IsRootElementLoaded) {
                    string loadedXml = ExcelDocument.ReadOpenXmlPartText(part);
                    if (loadedNative.Connections != null) {
                        result.Add(new WorkbookConnectionRoot(part, loadedNative.Connections));
                    }
                    budget?.ConsumeXmlBytes(loadedXml.Length);
                    continue;
                }

                string xml = ReadWorkbookConnectionXml(part, budget);
                if (string.IsNullOrWhiteSpace(xml)) continue;
                try {
                    var connections = new Connections(xml);
                    if (part is ConnectionsPart nativePart) {
                        nativePart.Connections = connections;
                    }
                    result.Add(new WorkbookConnectionRoot(part, connections));
                } catch (Exception exception) when (
                    exception is ArgumentException
                    || exception is InvalidDataException
                    || exception is XmlException) {
                    throw new InvalidDataException("Workbook connection metadata is not valid SpreadsheetML connections XML.", exception);
                }
            }
            return result;
        }

        private static string ReadWorkbookConnectionXml(
            OpenXmlPart part,
            MutationPlanScanBudget? budget) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(
                stream,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 4096,
                leaveOpen: false);
            var text = new StringBuilder(Math.Min(8192, ExcelDocument.MaximumWorkbookConnectionMetadataCharacters));
            var buffer = new char[4096];
            while (true) {
                int requested = budget?.GetBudgetedReadSize(buffer.Length) ?? buffer.Length;
                int read = reader.Read(buffer, 0, requested);
                if (read == 0) return text.ToString();
                budget?.ConsumeXmlBytes(read);
                if (text.Length > ExcelDocument.MaximumWorkbookConnectionMetadataCharacters - read) {
                    throw new InvalidDataException(
                        $"Workbook connection metadata exceeds {ExcelDocument.MaximumWorkbookConnectionMetadataCharacters} characters.");
                }
                text.Append(buffer, 0, read);
            }
        }

    }
}
