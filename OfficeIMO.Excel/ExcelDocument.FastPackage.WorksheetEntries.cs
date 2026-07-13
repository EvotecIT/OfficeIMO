using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {

        private static void WriteWorksheetEntry(ZipArchive archive, FastWorksheetPackageModel model) {
            var entry = archive.CreateEntry(model.WorksheetPath, CompressionLevel.Fastest);
            var worksheet = model.Worksheet;
            string dimension = worksheet.SheetDimension?.Reference?.Value ?? ExcelSheet.ComputeSheetDimensionReference(worksheet);
            var builder = new System.Text.StringBuilder(4096);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, Utf8NoBom);

            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            if (model.RequiresRelationshipNamespace) {
                builder.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
            }

            builder.Append(">");
            WriteBuilderAndClear(writer, builder);
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetProperties>());

            builder.Append("<dimension ref=\"");
            AppendXmlEscaped(builder, dimension);
            builder.Append("\"/>");
            WriteBuilderAndClear(writer, builder);

            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetViews>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetFormatProperties>());

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns != null) {
                AppendColumns(builder, columns);
                WriteBuilderAndClear(writer, builder);
            }

            writer.Write("<sheetData>");

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) {
                foreach (var row in sheetData.Elements<Row>()) {
                    AppendSimpleRowStart(builder, row);

                    foreach (var cell in row.Elements<Cell>()) {
                        AppendSimpleCell(builder, cell);
                    }

                    builder.Append("</row>");
                    WriteBuilderAndClear(writer, builder);
                }
            }

            writer.Write("</sheetData>");
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetCalculationProperties>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetProtection>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<Scenarios>());

            var autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter != null) {
                writer.Write(autoFilter.OuterXml);
            }

            WriteOptionalElement(writer, worksheet.GetFirstChild<SortState>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<MergeCells>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PhoneticProperties>());
            WriteOptionalElements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>(writer, worksheet);
            WriteOptionalElement(writer, worksheet.GetFirstChild<DataValidations>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<Hyperlinks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PrintOptions>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PageMargins>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PageSetup>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<HeaderFooter>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<RowBreaks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<ColumnBreaks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<CellWatches>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>());

            var tableParts = worksheet.GetFirstChild<TableParts>();
            if (tableParts != null && model.TablePartPaths.Count > 0) {
                builder.Append("<tableParts count=\"");
                AppendInvariant(builder, model.TablePartPaths.Count);
                builder.Append("\">");
                foreach (var tablePart in tableParts.Elements<TablePart>()) {
                    string? id = tablePart.Id?.Value;
                    if (id == null || !model.TablePartPaths.ContainsKey(id)) {
                        continue;
                    }

                    builder.Append("<tablePart r:id=\"");
                    AppendXmlEscaped(builder, id);
                    builder.Append("\"/>");
                }

                builder.Append("</tableParts>");
                WriteBuilderAndClear(writer, builder);
            }

            writer.Write("</worksheet>");
        }

        private static void WriteBuilderAndClear(StreamWriter writer, System.Text.StringBuilder builder) {
            if (builder.Length == 0) {
                return;
            }

#if NET6_0_OR_GREATER
            writer.Write(builder);
#else
            writer.Write(builder.ToString());
#endif
            builder.Clear();
            if (builder.Capacity > 65536) {
                builder.Capacity = 4096;
            }
        }

        private static void WriteOptionalElement(StreamWriter writer, OpenXmlElement? element) {
            if (element != null) {
                writer.Write(element.OuterXml);
            }
        }

        private static void WriteOptionalElements<TElement>(StreamWriter writer, OpenXmlElement parent)
            where TElement : OpenXmlElement {
            foreach (var element in parent.Elements<TElement>()) {
                writer.Write(element.OuterXml);
            }
        }

        private static void AppendColumns(System.Text.StringBuilder builder, Columns columns) {
            builder.Append("<cols>");
            foreach (var column in columns.Elements<Column>()) {
                builder.Append("<col");
                AppendUIntAttribute(builder, "min", column.Min);
                AppendUIntAttribute(builder, "max", column.Max);
                if (column.Width != null) {
                    builder.Append(" width=\"");
                    builder.Append(InvariantNumberText.Get(column.Width.Value));
                    builder.Append('"');
                }

                AppendBooleanAttribute(builder, "bestFit", column.BestFit);
                AppendBooleanAttribute(builder, "customWidth", column.CustomWidth);
                AppendBooleanAttribute(builder, "hidden", column.Hidden);
                AppendUIntAttribute(builder, "style", column.Style);
                AppendByteAttribute(builder, "outlineLevel", column.OutlineLevel);
                AppendBooleanAttribute(builder, "collapsed", column.Collapsed);
                AppendBooleanAttribute(builder, "phonetic", column.Phonetic);
                builder.Append("/>");
            }

            builder.Append("</cols>");
        }

        private static void AppendSimpleRowStart(System.Text.StringBuilder builder, Row row) {
            builder.Append("<row");
            AppendUIntAttribute(builder, "r", row.RowIndex);
            AppendBooleanAttribute(builder, "hidden", row.Hidden);
            if (row.Height != null) {
                builder.Append(" ht=\"");
                builder.Append(InvariantNumberText.Get(row.Height.Value));
                builder.Append('"');
            }

            AppendBooleanAttribute(builder, "customHeight", row.CustomHeight);
            AppendByteAttribute(builder, "outlineLevel", row.OutlineLevel);
            AppendBooleanAttribute(builder, "collapsed", row.Collapsed);
            builder.Append('>');
        }

        private static void AppendSimpleCell(System.Text.StringBuilder builder, Cell cell) {
            string? text = cell.CellValue?.Text;
            var dataType = cell.DataType?.Value;

            builder.Append("<c");
            if (cell.CellReference != null) {
                builder.Append(" r=\"");
                AppendXmlEscaped(builder, cell.CellReference.Value ?? string.Empty);
                builder.Append('"');
            }

            if (cell.StyleIndex != null) {
                builder.Append(" s=\"");
                AppendInvariant(builder, cell.StyleIndex.Value);
                builder.Append('"');
            }

            if (dataType == CellValues.Number) {
                builder.Append(" t=\"n\"");
            } else if (dataType == CellValues.SharedString) {
                builder.Append(" t=\"s\"");
            } else if (dataType == CellValues.InlineString || cell.InlineString != null) {
                builder.Append(" t=\"inlineStr\"");
            } else if (dataType == CellValues.String) {
                builder.Append(" t=\"str\"");
            } else if (dataType == CellValues.Boolean) {
                builder.Append(" t=\"b\"");
            } else if (dataType == CellValues.Error) {
                builder.Append(" t=\"e\"");
            } else if (dataType == CellValues.Date) {
                builder.Append(" t=\"d\"");
            }

            builder.Append('>');
            if (cell.CellFormula != null) {
                AppendCellFormula(builder, cell.CellFormula);
            }

            if (cell.InlineString != null) {
                builder.Append(cell.InlineString.OuterXml);
                builder.Append("</c>");
                return;
            }

            if (cell.CellValue != null) {
                string valueText = text ?? string.Empty;
                if (valueText.Length == 0) {
                    builder.Append("<v/>");
                } else {
                    builder.Append("<v>");
                    AppendXmlEscaped(builder, valueText);
                    builder.Append("</v>");
                }
            }

            builder.Append("</c>");
        }

        private static void AppendCellFormula(System.Text.StringBuilder builder, CellFormula formula) {
            if (formula.HasAttributes) {
                builder.Append(formula.OuterXml);
                return;
            }

            builder.Append("<f>");
            AppendXmlEscaped(builder, formula.Text ?? string.Empty);
            builder.Append("</f>");
        }

        private static void WriteTextEntry(ZipArchive archive, string path, string text) {
            var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, Utf8NoBom);
            writer.Write(text);
        }

        private static void AppendUIntAttribute(System.Text.StringBuilder builder, string name, UInt32Value? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            AppendInvariant(builder, value.Value);
            builder.Append('"');
        }

        private static void AppendByteAttribute(System.Text.StringBuilder builder, string name, ByteValue? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            AppendInvariant(builder, value.Value);
            builder.Append('"');
        }

        private static void AppendBooleanAttribute(System.Text.StringBuilder builder, string name, BooleanValue? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            builder.Append(value.Value ? '1' : '0');
            builder.Append('"');
        }

        private static void AppendInvariant(System.Text.StringBuilder builder, int value)
            => builder.Append(InvariantNumberText.Get(value));

        private static void AppendInvariant(System.Text.StringBuilder builder, uint value)
            => builder.Append(InvariantNumberText.Get(value));

        private static XmlWriter CreateFastXmlWriter(Stream stream) =>
            XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = Utf8NoBom,
                CloseOutput = false,
                Indent = false,
                OmitXmlDeclaration = false
            });

        private static void AppendXmlEscaped(System.Text.StringBuilder builder, string text) {
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                switch (ch) {
                    case '&':
                        builder.Append("&amp;");
                        break;
                    case '<':
                        builder.Append("&lt;");
                        break;
                    case '>':
                        builder.Append("&gt;");
                        break;
                    case '"':
                        builder.Append("&quot;");
                        break;
                    case '\'':
                        builder.Append("&apos;");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }
        }
    }
}
