using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_ImportsPhase4CommentRichTextRuns() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4RichTextCommentWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsComment comment = Assert.Single(Assert.Single(legacy.Worksheets).Comments);
            Assert.Equal("Bold plain", comment.Text);
            Assert.Collection(
                comment.FormattingRuns,
                run => {
                    Assert.Equal(0, run.StartCharacter);
                    Assert.Equal(5, run.FontIndex);
                },
                run => {
                    Assert.Equal(4, run.StartCharacter);
                    Assert.Equal(0, run.FontIndex);
                });

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            ExcelCommentInfo projectedComment = Assert.Single(document.Sheets[0].GetComments());
            Assert.Equal("Bold plain", projectedComment.Text);
            Assert.Collection(
                projectedComment.RichTextRuns,
                run => {
                    Assert.Equal("Bold", run.Text);
                    Assert.Equal("Consolas", run.FontName);
                    Assert.Equal(13d, run.FontSize);
                    Assert.Equal("FF123456", run.FontColor);
                    Assert.True(run.Bold);
                    Assert.True(run.Italic);
                    Assert.True(run.Underline);
                },
                run => {
                    Assert.Equal(" plain", run.Text);
                    Assert.Equal("Arial", run.FontName);
                    Assert.Equal(11d, run.FontSize);
                    Assert.False(run.Bold);
                });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetCommentsPart commentsPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().WorksheetCommentsPart!;
            DocumentFormat.OpenXml.Spreadsheet.Comment openXmlComment = Assert.Single(commentsPart.Comments!.CommentList!.Elements<DocumentFormat.OpenXml.Spreadsheet.Comment>());
            List<Run> openXmlRuns = openXmlComment.CommentText!.Elements<Run>().ToList();
            Assert.Equal(2, openXmlRuns.Count);
            Assert.Equal("Bold", openXmlRuns[0].Text!.Text);
            Assert.NotNull(openXmlRuns[0].RunProperties!.GetFirstChild<Bold>());
            Assert.Equal("Consolas", openXmlRuns[0].RunProperties!.GetFirstChild<RunFont>()!.Val!.Value);
            Assert.Equal(" plain", openXmlRuns[1].Text!.Text);
        }

        [Fact]
        public void LegacyXls_Load_PreservesCommentAnchorGeometry() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AnchoredCommentWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            LegacyXlsComment comment = Assert.Single(Assert.Single(legacy.Worksheets).Comments);
            Assert.True(comment.HasAnchor);
            LegacyXlsDrawingAnchor anchor = comment.Anchor!;
            Assert.Equal((ushort)0x0000, anchor.Flags);
            Assert.Equal((ushort)1, anchor.StartColumn);
            Assert.Equal((ushort)10, anchor.StartDx);
            Assert.Equal((ushort)2, anchor.StartRow);
            Assert.Equal((ushort)20, anchor.StartDy);
            Assert.Equal((ushort)3, anchor.EndColumn);
            Assert.Equal((ushort)30, anchor.EndDx);
            Assert.Equal((ushort)4, anchor.EndRow);
            Assert.Equal((ushort)40, anchor.EndDy);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.CommentsByAnchorRange["R2C1:R4C3"]);
            Assert.Equal(1, report.CommentsByAnchorOffset["StartDx:10;StartDy:20;EndDx:30;EndDy:40"]);
            Assert.Equal(1, report.CommentsByAnchorFlags["Flags:0x0000"]);
            Assert.Contains("Comments By Anchor Range", report.ToMarkdown());

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            VmlDrawingPart vmlPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.Single().VmlDrawingParts);
            using var reader = new StreamReader(vmlPart.GetStream());
            string vml = reader.ReadToEnd();
            Assert.Contains("<x:Anchor>1, 10, 2, 20, 3, 30, 4, 40</x:Anchor>", vml, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsVisibleCommentsAsVisibleNotes() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5VisibleCommentWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            VmlDrawingPart vmlPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.Single().VmlDrawingParts);
            using var reader = new StreamReader(vmlPart.GetStream());
            string vml = reader.ReadToEnd();
            Assert.Contains("visibility:visible", vml, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<x:Visible", vml, StringComparison.OrdinalIgnoreCase);
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreatePhase4AnchoredCommentWorkbookStream() {
                const string text = "Anchored note";
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "AnchoredComment"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Review me"));
                WriteRecord(stream, 0x00ec, BuildDrawingWithPictureShapePayload());
                WriteRecord(stream, 0x005d, BuildNoteObjectPayload(1));
                WriteRecord(stream, 0x01b6, BuildTxoPayload(text, 8));
                WriteRecord(stream, 0x003c, BuildCompressedUnicodeStringNoCchPayload(text));
                WriteRecord(stream, 0x003c, BuildTxoRunsPayload(((ushort)text.Length, 0)));
                WriteRecord(stream, 0x001c, BuildNotePayload(0, 0, 1, "Legacy Author"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4RichTextCommentWorkbookStream() {
                const string text = "Bold plain";
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "RichComment"));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: false, italic: false, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: true, italic: false, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: false, italic: true, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: true, italic: true, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Consolas", 13d, bold: true, italic: true, underline: true, colorIndex: 0x0008));
                WriteRecord(stream, 0x0092, BuildPalettePayload("FF123456"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Review me"));
                WriteRecord(stream, 0x005d, BuildNoteObjectPayload(1));
                WriteRecord(stream, 0x01b6, BuildTxoPayload(text, 24));
                WriteRecord(stream, 0x003c, BuildCompressedUnicodeStringNoCchPayload(text));
                WriteRecord(stream, 0x003c, BuildTxoRunsPayload((0, 5), (4, 0), ((ushort)text.Length, 0)));
                WriteRecord(stream, 0x001c, BuildNotePayload(0, 0, 1, "Legacy Author"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5VisibleCommentWorkbookStream() {
                const string text = "Visible note";
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "VisibleComment"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Review me"));
                WriteRecord(stream, 0x005d, BuildNoteObjectPayload(1));
                WriteRecord(stream, 0x01b6, BuildTxoPayload(text, 8));
                WriteRecord(stream, 0x003c, BuildCompressedUnicodeStringNoCchPayload(text));
                WriteRecord(stream, 0x003c, BuildTxoRunsPayload((ushort)text.Length));
                WriteRecord(stream, 0x001c, BuildNotePayload(0, 0, 1, "Legacy Author", visible: true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildTxoPayload(string text, ushort formattingRunBytes) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0212);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, checked((ushort)text.Length));
                WriteUInt16(stream, formattingRunBytes);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildTxoRunsPayload(params (ushort StartCharacter, ushort FontIndex)[] runs) {
                using var stream = new MemoryStream();
                foreach ((ushort startCharacter, ushort fontIndex) in runs) {
                    WriteUInt16(stream, startCharacter);
                    WriteUInt16(stream, fontIndex);
                    WriteUInt32(stream, 0);
                }

                return stream.ToArray();
            }
        }
    }
}
