using System.Text;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static partial class LegacyXlsTestWorkbookBuilder {
            private static readonly byte[] UrlMonikerClsid = {
                0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11,
                0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b
            };

            private static readonly byte[] FileMonikerClsid = {
                0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
            };

            internal static byte[] CreatePhase4HyperlinkWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Links"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "OfficeIMO"));
                WriteRecord(stream, 0x01b8, BuildExternalUrlHLinkPayload(0, 0, 0, 0, "https://officeimo.net/legacy-xls", "OfficeIMO"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4InternalHyperlinkWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long linksBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Links"));
                long targetBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Target"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int linksSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Jump"));
                WriteRecord(stream, 0x01b8, BuildInternalLocationHLinkPayload(0, 0, 0, 0, "'Target'!B2", "Jump"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int targetSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 1, "Destination"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(linksSheetOffset), 0, bytes, checked((int)linksBoundSheetPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(targetSheetOffset), 0, bytes, checked((int)targetBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4FileHyperlinkWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Files"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Budget"));
                WriteRecord(stream, 0x01b8, BuildFileHLinkPayload(0, 0, 0, 0, @"C:\Data\Budget.pdf", "Budget"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4UncFileHyperlinkWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "UncLinks"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Shared Budget"));
                WriteRecord(stream, 0x01b8, BuildFileHLinkPayload(0, 0, 0, 0, @"\\fileserver\share\Budget.pdf", "Shared Budget"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4RelativeFileHyperlinkWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "RelativeLinks"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Relative Budget"));
                WriteRecord(stream, 0x01b8, BuildFileHLinkPayload(0, 0, 0, 0, @"..\Docs\Budget.pdf", "Relative Budget", parentDirectoryCount: 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4CommentWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Comments"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Review me"));
                WriteRecord(stream, 0x005d, BuildNoteObjectPayload(1));
                WriteRecord(stream, 0x01b6, BuildTxoPayload("Imported legacy note"));
                WriteRecord(stream, 0x003c, BuildCompressedUnicodeStringNoCchPayload("Imported legacy note"));
                WriteRecord(stream, 0x003c, BuildTxoRunsPayload((ushort)"Imported legacy note".Length));
                WriteRecord(stream, 0x001c, BuildNotePayload(0, 0, 1, "Legacy Author"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4WorksheetProtectionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Protected"));
                WriteRecord(stream, 0x0012, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x0013, BuildUInt16Payload(0xcafe));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Locked"));
                WriteRecord(stream, 0x0012, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x0013, BuildUInt16Payload(0xbeef));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateCalculationSettingsWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Calc"));
                WriteRecord(stream, 0x000d, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x000c, BuildUInt16Payload(42));
                WriteRecord(stream, 0x000e, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x000f, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x0010, BuildDoublePayload(0.001d));
                WriteRecord(stream, 0x0011, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x005f, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Calculated"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateWorksheetMetadataWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Metadata"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 1, "Active"));
                WriteRecord(stream, 0x0081, BuildUInt16Payload(0xf1e1));
                WriteRecord(stream, 0x0080, BuildGutsPayload(rowOutlineLevelRaw: 3, columnOutlineLevelRaw: 4));
                WriteRecord(stream, 0x0082, BuildUInt16Payload(1));
                WriteRecord(stream, 0x020b, BuildIndexPayload(firstRow: 0, rowAfterLast: 4, reservedRecordOffset: 1234, dbCellOffsets: new uint[] { 200, 240 }));
                WriteRecord(stream, 0x001d, BuildSelectionPayload());
                WriteRecord(stream, 0x0090, BuildSortPayload("Region", "Amount", "Date"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4PrintPageSetupWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Print"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Printable"));
                WriteRecord(stream, 0x0026, BuildDoublePayload(0.25d));
                WriteRecord(stream, 0x0027, BuildDoublePayload(0.35d));
                WriteRecord(stream, 0x0028, BuildDoublePayload(0.5d));
                WriteRecord(stream, 0x0029, BuildDoublePayload(0.6d));
                WriteRecord(stream, 0x00a1, BuildSetupPayload(scale: 125, fitToWidth: 1, fitToHeight: 2, landscape: true, header: 0.4d, footer: 0.45d));
                WriteRecord(stream, 0x004d, BuildPrinterSettingsPayload());
                WriteRecord(stream, 0x0033, BuildUInt16Payload(2));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4PrintOptionsWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "PrintOptions"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Printable"));
                WriteRecord(stream, 0x002a, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x002b, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x0083, BuildUInt16Payload(0x0001));
                WriteRecord(stream, 0x0084, BuildUInt16Payload(0x0000));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4PageBreaksWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Breaks"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Breaks"));
                WriteRecord(stream, 0x001b, BuildHorizontalPageBreaksPayload((3, 0, 255)));
                WriteRecord(stream, 0x001a, BuildVerticalPageBreaksPayload((2, 0, 20)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4ZoomScaleWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Zoom"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Zoomed"));
                WriteRecord(stream, 0x00a0, BuildZoomScalePayload(3, 2));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4HeaderFooterWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "HeaderFooter"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Printable"));
                WriteRecord(stream, 0x0014, BuildUnicodeStringPayload("&LLeft &P&L&E Again&CQuarterly&RConfidential"));
                WriteRecord(stream, 0x0015, BuildUnicodeStringPayload("&CPage &P of &N"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4DefinedNamesWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Names"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload("DataRange", BuildNameArea3dFormula(0, 0, 0, 1, 1), localSheetIndex: 0, hidden: false, builtIn: false));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload("HiddenCell", BuildNameRef3dFormula(0, 2, 2), localSheetIndex: 1, hidden: true, builtIn: false));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload(((char)0x06).ToString(), BuildNameArea3dFormula(0, 0, 0, 3, 1), localSheetIndex: 1, hidden: false, builtIn: true));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload(((char)0x07).ToString(), BuildNamePrintTitlesFormula(0), localSheetIndex: 1, hidden: false, builtIn: true));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload(((char)0x0d).ToString(), BuildNameArea3dFormula(0, 0, 0, 2, 1), localSheetIndex: 1, hidden: true, builtIn: true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Name"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 1, "Value"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 0, "OfficeIMO"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 1, "42"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 0, "Legacy"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 1, "17"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 2, "Hidden"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4AutoFilterCriteriaWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Filtered"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload(((char)0x0d).ToString(), BuildNameArea3dFormula(0, 0, 0, 3, 1), localSheetIndex: 1, hidden: true, builtIn: true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Status"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 1, "Amount"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 0, "Open"));
                WriteRecord(stream, 0x027e, BuildRkPayload(1, 1, 0, EncodeRkInteger(5)));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 0, "Closed"));
                WriteRecord(stream, 0x027e, BuildRkPayload(2, 1, 0, EncodeRkInteger(15)));
                WriteRecord(stream, 0x0204, BuildLabelPayload(3, 0, "Open"));
                WriteRecord(stream, 0x027e, BuildRkPayload(3, 1, 0, EncodeRkInteger(25)));
                WriteRecord(stream, 0x009d, BuildAutoFilterInfoPayload(2));
                WriteRecord(stream, 0x009b, Array.Empty<byte>());
                WriteRecord(stream, 0x009e, BuildAutoFilterStringEqualsPayload(0, "Open"));
                WriteRecord(stream, 0x009e, BuildAutoFilterNumberGreaterThanOrEqualPayload(1, 10d));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4AutoFilterTop10WorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "TopScores"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload(((char)0x0d).ToString(), BuildNameArea3dFormula(0, 0, 0, 4, 0), localSheetIndex: 1, hidden: true, builtIn: true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Score"));
                WriteRecord(stream, 0x027e, BuildRkPayload(1, 0, 0, EncodeRkInteger(10)));
                WriteRecord(stream, 0x027e, BuildRkPayload(2, 0, 0, EncodeRkInteger(20)));
                WriteRecord(stream, 0x027e, BuildRkPayload(3, 0, 0, EncodeRkInteger(30)));
                WriteRecord(stream, 0x027e, BuildRkPayload(4, 0, 0, EncodeRkInteger(40)));
                WriteRecord(stream, 0x009d, BuildAutoFilterInfoPayload(1));
                WriteRecord(stream, 0x009b, Array.Empty<byte>());
                WriteRecord(stream, 0x009e, BuildAutoFilterTop10Payload(0, 10, isTop: true, isPercent: false));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4AutoFilterBlankNonBlankWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "BlankFilters"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload(((char)0x0d).ToString(), BuildNameArea3dFormula(0, 0, 0, 4, 1), localSheetIndex: 1, hidden: true, builtIn: true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "MaybeBlank"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 1, "MaybeValue"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 0, "Open"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 1, "North"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(3, 0, "Closed"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(4, 1, "South"));
                WriteRecord(stream, 0x009d, BuildAutoFilterInfoPayload(2));
                WriteRecord(stream, 0x009b, Array.Empty<byte>());
                WriteRecord(stream, 0x009e, BuildAutoFilterBlanksPayload(0));
                WriteRecord(stream, 0x009e, BuildAutoFilterNonBlanksPayload(1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5UnsupportedSheetTypesWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Data"));
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Macro1", sheetType: 0x01));
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Chart1", sheetType: 0x02));
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Module1", sheetType: 0x06));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Imported"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)dataBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5ChartSheetSubstreamWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Data"));
                long chartBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ChartOnly", sheetType: 0x02));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int dataSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Imported"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int chartSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x20, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x1001, Array.Empty<byte>());
                WriteRecord(stream, 0x1002, BuildChartPayload(100, 200, 3000, 2200));
                WriteRecord(stream, 0x1014, Array.Empty<byte>());
                WriteRecord(stream, 0x101d, BuildAxisPayload(0x0001));
                WriteRecord(stream, 0x1045, BuildUInt16Payload(1));
                WriteRecord(stream, 0x1003, BuildSeriesPayload(0x0003, categoryCount: 4, valueCount: 4, bubbleSizeCount: 0));
                WriteRecord(stream, 0x1006, BuildDataFormatPayload(pointIndex: 0xffff, seriesIndex: 2, order: 1));
                WriteRecord(stream, 0x1007, BuildLineFormatPayload(style: 0x0001, weight: 1, flags: 0x0004, colorIndex: 0x004d));
                WriteRecord(stream, 0x100a, BuildAreaFormatPayload(pattern: 0x0001, flags: 0x0003, foregroundColorIndex: 0x004e, backgroundColorIndex: 0x004d));
                WriteRecord(stream, 0x1009, BuildMarkerFormatPayload(markerType: 0x0008, flags: 0x0021, foregroundColorIndex: 0x004e, backgroundColorIndex: 0x004d, sizeTwips: 240));
                WriteRecord(stream, 0x101b, Array.Empty<byte>());
                WriteRecord(stream, 0x01b6, new byte[18]);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(dataSheetOffset), 0, bytes, checked((int)dataBoundSheetPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(chartSheetOffset), 0, bytes, checked((int)chartBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5DialogSheetWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dialogBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Dialog1"));
                long dataBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Data"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int dialogSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0081, BuildUInt16Payload(0x0010));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "DialogOnly"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int dataSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0081, BuildUInt16Payload(0x0000));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Imported"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(dialogSheetOffset), 0, bytes, checked((int)dialogBoundSheetPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(dataSheetOffset), 0, bytes, checked((int)dataBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5ExternalReferencesWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Data"));
                WriteRecord(stream, 0x01ae, BuildSupBookExternalWorkbookPayload("C:\\Data\\Budget.xls", "Jan", "Feb"));
                WriteRecord(stream, 0x0023, BuildExternalNamePayload("TaxRate"));
                WriteRecord(stream, 0x01b7, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Imported"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)dataBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5PreserveOnlyFeatureDetailsWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "FeatureMap"));
                WriteRecord(stream, 0x00eb, BuildEscherHeaderPayload(0xf000, instance: 2, version: 0x0f, length: 8));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Imported"));
                WriteRecord(stream, 0x005d, BuildObjectPayload(0x0008, 1));
                WriteRecord(stream, 0x00ec, BuildEscherHeaderPayload(0xf002, instance: 1, version: 0x0f, length: 0));
                WriteRecord(stream, 0x1002, Array.Empty<byte>());
                WriteRecord(stream, 0x00b0, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)dataBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5PivotTableMetadataWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "PivotMeta"));
                WriteRecord(stream, 0x00b0, Array.Empty<byte>());
                WriteRecord(stream, 0x00c1, BuildSxdiPayload());
                WriteRecord(stream, 0x00c5, Array.Empty<byte>());
                WriteRecord(stream, 0x00c8, Array.Empty<byte>());
                WriteRecord(stream, 0x00cf, Array.Empty<byte>());
                WriteRecord(stream, 0x00f1, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Pivot data"));
                WriteRecord(stream, 0x00b1, Array.Empty<byte>());
                WriteRecord(stream, 0x00b2, Array.Empty<byte>());
                WriteRecord(stream, 0x00d7, BuildSxRngPayload());
                WriteRecord(stream, 0x00f9, Array.Empty<byte>());
                WriteRecord(stream, 0x00ff, BuildSxVdExPayload());
                WriteRecord(stream, 0x0858, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)dataBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5DataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Validation"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Validated"));
                WriteRecord(stream, 0x01b2, Array.Empty<byte>());
                WriteRecord(stream, 0x01be, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4DataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Validation"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Validated"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildWholeNumberDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4DecimalDataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "DecimalValidation"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Discount"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildDecimalDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4ListDataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ListValidation"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Status"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildListDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4RangeListDataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "RangeListValidation"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Open"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 0, "Closed"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 0, "Pending"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 7, "Status"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildRangeListDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4CrossSheetRangeListDataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long optionsBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Options"));
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "CrossSheetValidation"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int optionsSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Open"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 0, "Closed"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 0, "Pending"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int validationSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 7, "Status"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildCrossSheetRangeListDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(optionsSheetOffset), 0, bytes, checked((int)optionsBoundSheetPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(validationSheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4NamedListDataValidationWorkbookStream() {
                return CreateNamedListDataValidationWorkbookStream("NamedListValidation", localSheetIndex: 0);
            }

            internal static byte[] CreatePhase4SheetLocalNamedListDataValidationWorkbookStream() {
                return CreateNamedListDataValidationWorkbookStream("LocalNamedListValidation", localSheetIndex: 1);
            }

            private static byte[] CreateNamedListDataValidationWorkbookStream(string sheetName, ushort localSheetIndex) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, sheetName));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload("StatusOptions", BuildNameArea3dFormula(0, 0, 0, 2, 0), localSheetIndex, hidden: false, builtIn: false));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Open"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 0, "Closed"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 0, "Pending"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 7, "Status"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildNamedListDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4TypedDataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "TypedValidation"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 4, "Date"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 5, "Time"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 6, "Text length"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(3));
                WriteRecord(stream, 0x01be, BuildDateDataValidationPayload());
                WriteRecord(stream, 0x01be, BuildTimeDataValidationPayload());
                WriteRecord(stream, 0x01be, BuildTextLengthDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4CustomFormulaDataValidationWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long validationBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "CustomValidation"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Left"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 1, "Right"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 5, "Total"));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildCustomFormulaDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validationBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4ConditionalFormattingWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long formattingBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ConditionalRule"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Score"));
                WriteRecord(stream, 0x027e, BuildRkPayload(1, 0, 0, EncodeRkInteger(5)));
                WriteRecord(stream, 0x027e, BuildRkPayload(2, 0, 0, EncodeRkInteger(15)));
                WriteRecord(stream, 0x01b0, BuildConditionalFormattingRangePayload(0, 0, 2, 0, 1));
                WriteRecord(stream, 0x01b1, BuildCellIsGreaterThanConditionalFormattingRulePayload(10));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)formattingBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4ConditionalFormulaFormattingWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long formattingBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ConditionalFormula"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Score"));
                WriteRecord(stream, 0x027e, BuildRkPayload(1, 0, 0, EncodeRkInteger(5)));
                WriteRecord(stream, 0x027e, BuildRkPayload(2, 0, 0, EncodeRkInteger(15)));
                WriteRecord(stream, 0x01b0, BuildConditionalFormattingRangePayload(0, 0, 2, 0, 1));
                WriteRecord(stream, 0x01b1, BuildFormulaConditionalFormattingRulePayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)formattingBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5ConditionalFormattingExtensionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long formattingBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ConditionalExt"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Score"));
                WriteRecord(stream, 0x027e, BuildRkPayload(1, 0, 0, EncodeRkInteger(5)));
                WriteRecord(stream, 0x027e, BuildRkPayload(2, 0, 0, EncodeRkInteger(15)));
                WriteRecord(stream, 0x01b0, BuildConditionalFormattingRangePayload(0, 0, 2, 0, 1, headerId: 2));
                WriteRecord(stream, 0x01b1, BuildCellIsGreaterThanConditionalFormattingRulePayload(10));
                WriteRecord(stream, 0x087b, BuildConditionalFormattingExtensionPayload(headerId: 1, priority: 7, stopIfTrue: true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)formattingBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5EmbeddedChartBeforeWorksheetFeaturesWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "NestedFeatures"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Status"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 1, "Amount"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 0, "Open"));
                WriteRecord(stream, 0x027e, BuildRkPayload(1, 1, 0, EncodeRkInteger(125)));
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x20, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x1002, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                WriteRecord(stream, 0x01b0, BuildConditionalFormattingRangePayload(1, 1, 4, 1, 1));
                WriteRecord(stream, 0x01b1, BuildCellIsGreaterThanConditionalFormattingRulePayload(100));
                WriteRecord(stream, 0x01b2, BuildDataValidationCollectionPayload(1));
                WriteRecord(stream, 0x01be, BuildExcelInlineListDataValidationPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)dataBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateExcelInlineListDataValidationPayloadForTest() {
                return BuildExcelInlineListDataValidationPayload();
            }

            internal static byte[] CreatePhase5ConditionalFormattingWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long formattingBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Conditional"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Formatted"));
                WriteRecord(stream, 0x01b0, Array.Empty<byte>());
                WriteRecord(stream, 0x01b1, Array.Empty<byte>());
                WriteRecord(stream, 0x087a, Array.Empty<byte>());
                WriteRecord(stream, 0x087b, Array.Empty<byte>());
                WriteRecord(stream, 0x088d, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)formattingBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5DifferentialFormatWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long formattingBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Differential"));
                WriteRecord(stream, 0x088d, BuildDifferentialFormatBackgroundColorPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Formatted"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)formattingBoundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildDifferentialFormatBackgroundColorPayload() {
                return new byte[] {
                    0x8d, 0x08, 0x00, 0x00, 0x00, 0x00,
                    0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                    0x03, 0x00,
                    0x00, 0x00, 0x01, 0x00,
                    0x02, 0x00, 0x0c, 0x00,
                    0x05, 0xff, 0x00, 0x00, 0xff, 0xff, 0x00, 0xff
                };
            }

            private static byte[] BuildConditionalFormattingRangePayload(
                ushort firstRow,
                ushort firstColumn,
                ushort lastRow,
                ushort lastColumn,
                ushort ruleCount,
                ushort headerId = 0) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, ruleCount);
                WriteUInt16(stream, headerId);
                WriteCellRange(stream, firstRow, firstColumn, lastRow, lastColumn);
                WriteUInt16(stream, 1);
                WriteCellRange(stream, firstRow, firstColumn, lastRow, lastColumn);
                return stream.ToArray();
            }

            private static byte[] BuildCellIsGreaterThanConditionalFormattingRulePayload(ushort threshold) {
                byte[] formula = BuildPtgIntFormula(threshold);
                using var stream = new MemoryStream();
                stream.WriteByte(0x01);
                stream.WriteByte(0x05);
                WriteUInt16(stream, checked((ushort)formula.Length));
                WriteUInt16(stream, 0);
                stream.Write(formula, 0, formula.Length);
                return stream.ToArray();
            }

            private static byte[] BuildFormulaConditionalFormattingRulePayload() {
                byte[] formula = BuildPtgRefGreaterThanIntFormula(0, 0, 10);
                using var stream = new MemoryStream();
                stream.WriteByte(0x02);
                stream.WriteByte(0);
                WriteUInt16(stream, checked((ushort)formula.Length));
                WriteUInt16(stream, 0);
                stream.Write(formula, 0, formula.Length);
                return stream.ToArray();
            }

            private static byte[] BuildConditionalFormattingExtensionPayload(ushort headerId, ushort priority, bool stopIfTrue) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x087b);
                WriteUInt16(stream, headerId);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                stream.WriteByte(0x05);
                stream.WriteByte(0);
                WriteUInt16(stream, priority);
                stream.WriteByte((byte)(stopIfTrue ? 0x03 : 0x01));
                stream.WriteByte(0);
                stream.WriteByte(16);
                for (int i = 0; i < 16; i++) {
                    stream.WriteByte(0);
                }

                return stream.ToArray();
            }

            private static byte[] BuildPtgRefGreaterThanIntFormula(ushort row, ushort column, ushort threshold) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x24);
                WriteUInt16(stream, row);
                WriteUInt16(stream, unchecked((ushort)(0xc000 | column)));
                stream.WriteByte(0x1e);
                WriteUInt16(stream, threshold);
                stream.WriteByte(0x0d);
                return stream.ToArray();
            }

            private static void WriteCellRange(Stream stream, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
            }

            private static byte[] BuildExternalUrlHLinkPayload(
                ushort firstRow,
                ushort firstColumn,
                ushort lastRow,
                ushort lastColumn,
                string url,
                string? displayName) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, 2);
                uint flags = 0x00000001 | 0x00000002;
                if (!string.IsNullOrEmpty(displayName)) {
                    flags |= 0x00000010;
                }

                WriteUInt32(stream, flags);
                if (!string.IsNullOrEmpty(displayName)) {
                    WriteHyperlinkString(stream, displayName!);
                }

                stream.Write(UrlMonikerClsid, 0, UrlMonikerClsid.Length);
                byte[] urlBytes = Encoding.Unicode.GetBytes(url + '\0');
                WriteUInt32(stream, checked((uint)urlBytes.Length));
                stream.Write(urlBytes, 0, urlBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildInternalLocationHLinkPayload(
                ushort firstRow,
                ushort firstColumn,
                ushort lastRow,
                ushort lastColumn,
                string location,
                string? displayName) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, 2);
                uint flags = 0x00000008;
                if (!string.IsNullOrEmpty(displayName)) {
                    flags |= 0x00000010;
                }

                WriteUInt32(stream, flags);
                if (!string.IsNullOrEmpty(displayName)) {
                    WriteHyperlinkString(stream, displayName!);
                }

                WriteHyperlinkString(stream, location);
                return stream.ToArray();
            }

            private static byte[] BuildFileHLinkPayload(
                ushort firstRow,
                ushort firstColumn,
                ushort lastRow,
                ushort lastColumn,
                string path,
                string? displayName,
                ushort parentDirectoryCount = 0) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, 2);
                uint flags = 0x00000001 | 0x00000002;
                if (!string.IsNullOrEmpty(displayName)) {
                    flags |= 0x00000010;
                }

                WriteUInt32(stream, flags);
                if (!string.IsNullOrEmpty(displayName)) {
                    WriteHyperlinkString(stream, displayName!);
                }

                stream.Write(FileMonikerClsid, 0, FileMonikerClsid.Length);
                WriteUInt16(stream, parentDirectoryCount);
                byte[] pathBytes = Encoding.ASCII.GetBytes(path + '\0');
                WriteUInt32(stream, checked((uint)pathBytes.Length));
                stream.Write(pathBytes, 0, pathBytes.Length);
                WriteUInt16(stream, 0xffff);
                WriteUInt16(stream, 0xdead);
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildAutoFilterInfoPayload(ushort dropDownCount) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, dropDownCount);
                return stream.ToArray();
            }

            private static byte[] BuildAutoFilterStringEqualsPayload(ushort columnId, string value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, columnId);
                WriteUInt16(stream, 0x0004);
                WriteStringDoper(stream, value, LegacyAutoFilterComparisonEqual);
                WriteUnusedDoper(stream);
                WriteCompressedUnicodeStringNoCch(stream, value);
                return stream.ToArray();
            }

            private static byte[] BuildAutoFilterNumberGreaterThanOrEqualPayload(ushort columnId, double value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, columnId);
                WriteUInt16(stream, 0x0000);
                WriteNumberDoper(stream, value, LegacyAutoFilterComparisonGreaterThanOrEqual);
                WriteUnusedDoper(stream);
                return stream.ToArray();
            }

            private static byte[] BuildAutoFilterTop10Payload(ushort columnId, ushort value, bool isTop, bool isPercent) {
                using var stream = new MemoryStream();
                ushort flags = (ushort)(0x0010 | ((value & 0x01ff) << 7));
                if (isTop) {
                    flags |= 0x0020;
                }

                if (isPercent) {
                    flags |= 0x0040;
                }

                WriteUInt16(stream, columnId);
                WriteUInt16(stream, flags);
                WriteUnusedDoper(stream);
                WriteUnusedDoper(stream);
                return stream.ToArray();
            }

            private static byte[] BuildAutoFilterBlanksPayload(ushort columnId) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, columnId);
                WriteUInt16(stream, 0x0004);
                WriteBlankDoper(stream, LegacyAutoFilterComparisonEqual);
                WriteUnusedDoper(stream);
                return stream.ToArray();
            }

            private static byte[] BuildAutoFilterNonBlanksPayload(ushort columnId) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, columnId);
                WriteUInt16(stream, 0x0004);
                WriteNonBlankDoper(stream, LegacyAutoFilterComparisonNotEqual);
                WriteUnusedDoper(stream);
                return stream.ToArray();
            }

            private const byte LegacyAutoFilterComparisonEqual = 0x02;
            private const byte LegacyAutoFilterComparisonNotEqual = 0x05;
            private const byte LegacyAutoFilterComparisonGreaterThanOrEqual = 0x06;

            private static void WriteUnusedDoper(Stream stream) {
                stream.Write(new byte[10], 0, 10);
            }

            private static void WriteStringDoper(Stream stream, string value, byte comparison) {
                stream.WriteByte(0x06);
                stream.WriteByte(comparison);
                WriteUInt32(stream, 0);
                stream.WriteByte(checked((byte)value.Length));
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(0);
            }

            private static void WriteNumberDoper(Stream stream, double value, byte comparison) {
                stream.WriteByte(0x04);
                stream.WriteByte(comparison);
                byte[] valueBytes = BitConverter.GetBytes(value);
                stream.Write(valueBytes, 0, valueBytes.Length);
            }

            private static void WriteBlankDoper(Stream stream, byte comparison) {
                stream.WriteByte(0x0c);
                stream.WriteByte(comparison);
                stream.Write(new byte[8], 0, 8);
            }

            private static void WriteNonBlankDoper(Stream stream, byte comparison) {
                stream.WriteByte(0x0e);
                stream.WriteByte(comparison);
                stream.Write(new byte[8], 0, 8);
            }

            private static byte[] BuildNoteObjectPayload(ushort objectId) {
                return BuildObjectPayload(0x0019, objectId);
            }

            private static byte[] BuildObjectPayload(ushort objectType, ushort objectId) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0015);
                WriteUInt16(stream, 0x0012);
                WriteUInt16(stream, objectType);
                WriteUInt16(stream, objectId);
                WriteUInt16(stream, 0x4011);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, 0x000d);
                WriteUInt16(stream, 0x0016);
                stream.Write(new byte[16], 0, 16);
                WriteUInt16(stream, 0x0000);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildEscherHeaderPayload(ushort recordType, ushort instance, byte version, uint length) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)((instance << 4) | (version & 0x0f))));
                WriteUInt16(stream, recordType);
                WriteUInt32(stream, length);
                return stream.ToArray();
            }

            private static byte[] BuildChartPayload(int x, int y, int width, int height) {
                using var stream = new MemoryStream();
                WriteInt32(stream, x);
                WriteInt32(stream, y);
                WriteInt32(stream, width);
                WriteInt32(stream, height);
                return stream.ToArray();
            }

            private static byte[] BuildAxisPayload(ushort axisType) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, axisType);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildSeriesPayload(ushort categoryDataType, ushort categoryCount, ushort valueCount, ushort bubbleSizeCount) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, categoryDataType);
                WriteUInt16(stream, 0x0001);
                WriteUInt16(stream, categoryCount);
                WriteUInt16(stream, valueCount);
                WriteUInt16(stream, 0x0001);
                WriteUInt16(stream, bubbleSizeCount);
                return stream.ToArray();
            }

            private static byte[] BuildDataFormatPayload(ushort pointIndex, ushort seriesIndex, ushort order) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, pointIndex);
                WriteUInt16(stream, seriesIndex);
                WriteUInt16(stream, order);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildLineFormatPayload(ushort style, short weight, ushort flags, ushort colorIndex) {
                using var stream = new MemoryStream();
                WriteLongRgb(stream, 0x11, 0x22, 0x33);
                WriteUInt16(stream, style);
                WriteUInt16(stream, unchecked((ushort)weight));
                WriteUInt16(stream, flags);
                WriteUInt16(stream, colorIndex);
                return stream.ToArray();
            }

            private static byte[] BuildAreaFormatPayload(ushort pattern, ushort flags, ushort foregroundColorIndex, ushort backgroundColorIndex) {
                using var stream = new MemoryStream();
                WriteLongRgb(stream, 0xaa, 0xbb, 0xcc);
                WriteLongRgb(stream, 0x10, 0x20, 0x30);
                WriteUInt16(stream, pattern);
                WriteUInt16(stream, flags);
                WriteUInt16(stream, foregroundColorIndex);
                WriteUInt16(stream, backgroundColorIndex);
                return stream.ToArray();
            }

            private static byte[] BuildMarkerFormatPayload(ushort markerType, ushort flags, ushort foregroundColorIndex, ushort backgroundColorIndex, uint sizeTwips) {
                using var stream = new MemoryStream();
                WriteLongRgb(stream, 0xde, 0xad, 0xbe);
                WriteLongRgb(stream, 0x44, 0x55, 0x66);
                WriteUInt16(stream, markerType);
                WriteUInt16(stream, flags);
                WriteUInt16(stream, foregroundColorIndex);
                WriteUInt16(stream, backgroundColorIndex);
                WriteUInt32(stream, sizeTwips);
                return stream.ToArray();
            }

            private static void WriteLongRgb(Stream stream, byte red, byte green, byte blue) {
                stream.WriteByte(red);
                stream.WriteByte(green);
                stream.WriteByte(blue);
                stream.WriteByte(0);
            }

            private static byte[] BuildTxoPayload(string text) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0212);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, checked((ushort)text.Length));
                WriteUInt16(stream, 16);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildTxoRunsPayload(ushort textLength) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, textLength);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildNotePayload(ushort row, ushort column, ushort objectId, string author) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, objectId);
                WriteUInt16(stream, checked((ushort)author.Length));
                WriteCompressedUnicodeStringNoCch(stream, author);
                stream.WriteByte(0);
                return stream.ToArray();
            }

            private static byte[] BuildCompressedUnicodeStringNoCchPayload(string value) {
                using var stream = new MemoryStream();
                WriteCompressedUnicodeStringNoCch(stream, value);
                return stream.ToArray();
            }

            private static byte[] BuildZoomScalePayload(short numerator, short denominator) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, unchecked((ushort)numerator));
                WriteUInt16(stream, unchecked((ushort)denominator));
                return stream.ToArray();
            }

            private static byte[] BuildHorizontalPageBreaksPayload(params (ushort Row, ushort ColumnStart, ushort ColumnEnd)[] breaks) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)breaks.Length));
                foreach ((ushort row, ushort columnStart, ushort columnEnd) in breaks) {
                    WriteUInt16(stream, row);
                    WriteUInt16(stream, columnStart);
                    WriteUInt16(stream, columnEnd);
                }

                return stream.ToArray();
            }

            private static byte[] BuildVerticalPageBreaksPayload(params (ushort Column, ushort RowStart, ushort RowEnd)[] breaks) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)breaks.Length));
                foreach ((ushort column, ushort rowStart, ushort rowEnd) in breaks) {
                    WriteUInt16(stream, column);
                    WriteUInt16(stream, rowStart);
                    WriteUInt16(stream, rowEnd);
                }

                return stream.ToArray();
            }

            private static void WriteHyperlinkString(Stream stream, string value) {
                byte[] valueBytes = Encoding.Unicode.GetBytes(value + '\0');
                WriteUInt32(stream, checked((uint)((value.Length) + 1)));
                stream.Write(valueBytes, 0, valueBytes.Length);
            }

            private static byte[] BuildSupBookExternalWorkbookPayload(string target, params string[] sheetNames) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)sheetNames.Length));
                WriteUInt16(stream, checked((ushort)target.Length));
                WriteCompressedUnicodeStringNoCch(stream, target);
                foreach (string sheetName in sheetNames) {
                    WriteUInt16(stream, checked((ushort)sheetName.Length));
                    WriteCompressedUnicodeStringNoCch(stream, sheetName);
                }

                return stream.ToArray();
            }

            private static byte[] BuildDataValidationCollectionPayload(uint count) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0xffffffff);
                WriteUInt32(stream, count);
                return stream.ToArray();
            }

            private static byte[] BuildWholeNumberDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x01U
                    | 0x00000100U
                    | 0x00040000U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, "Age\0");
                WriteUnicodeString(stream, "Invalid age");
                WriteUnicodeString(stream, "Enter an age from 18 to 65.");
                WriteUnicodeString(stream, "Use a whole number between 18 and 65.");
                WriteDvFormula(stream, BuildPtgIntFormula(18));
                WriteDvFormula(stream, BuildPtgIntFormula(65));
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 3);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                return stream.ToArray();
            }

            private static byte[] BuildDecimalDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x02U
                    | (0x04U << 20)
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, string.Empty);
                WriteUnicodeString(stream, "Invalid discount");
                WriteUnicodeString(stream, string.Empty);
                WriteUnicodeString(stream, "Use a decimal value greater than 5.5.");
                WriteDvFormula(stream, BuildPtgNumFormula(5.5d));
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 3);
                WriteUInt16(stream, 2);
                WriteUInt16(stream, 2);
                return stream.ToArray();
            }

            private static byte[] BuildListDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x03U
                    | 0x00000080U
                    | 0x00000100U
                    | 0x00040000U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, "Status");
                WriteUnicodeString(stream, "Invalid status");
                WriteUnicodeString(stream, "Pick a status.");
                WriteUnicodeString(stream, "Use one of the listed statuses.");
                WriteDvFormula(stream, BuildPtgStrFormula("Open,Closed,Pending"));
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 3);
                WriteUInt16(stream, 3);
                return stream.ToArray();
            }

            private static byte[] BuildExcelInlineListDataValidationPayload() {
                return new byte[] {
                    0x83, 0x01, 0x0c, 0x00,
                    0x01, 0x00, 0x00, 0x00,
                    0x01, 0x00, 0x00, 0x00,
                    0x01, 0x00, 0x00, 0x00,
                    0x01, 0x00, 0x00, 0x00,
                    0x16, 0x00, 0x49, 0x00,
                    0x17, 0x13, 0x00, 0x4f, 0x70, 0x65, 0x6e, 0x00,
                    0x43, 0x6c, 0x6f, 0x73, 0x65, 0x64, 0x00,
                    0x50, 0x65, 0x6e, 0x64, 0x69, 0x6e, 0x67,
                    0x00, 0x00,
                    0x00, 0x00,
                    0x01, 0x00, 0x01, 0x00, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00
                };
            }

            private static byte[] BuildRangeListDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x03U
                    | 0x00000100U
                    | 0x00040000U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, "Status");
                WriteUnicodeString(stream, "Invalid status");
                WriteUnicodeString(stream, "Pick a status from the sheet range.");
                WriteUnicodeString(stream, "Use one of the status values from A1:A3.");
                WriteDvFormula(stream, BuildPtgAreaFormula(0, 0, 2, 0));
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 7);
                WriteUInt16(stream, 7);
                return stream.ToArray();
            }

            private static byte[] BuildCrossSheetRangeListDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x03U
                    | 0x00000100U
                    | 0x00040000U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, "Status");
                WriteUnicodeString(stream, "Invalid status");
                WriteUnicodeString(stream, "Pick a status from the Options sheet.");
                WriteUnicodeString(stream, "Use one of the status values from Options!A1:A3.");
                WriteDvFormula(stream, BuildPtgArea3dFormula(0, 0, 0, 2, 0));
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 7);
                WriteUInt16(stream, 7);
                return stream.ToArray();
            }

            private static byte[] BuildNamedListDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x03U
                    | 0x00000100U
                    | 0x00040000U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, "Status");
                WriteUnicodeString(stream, "Invalid status");
                WriteUnicodeString(stream, "Pick a status from the named range.");
                WriteUnicodeString(stream, "Use one of the values from StatusOptions.");
                WriteDvFormula(stream, BuildPtgNameFormula(1));
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 7);
                WriteUInt16(stream, 7);
                return stream.ToArray();
            }

            private static byte[] BuildDateDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x04U
                    | 0x00000100U
                    | 0x00040000U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, "Ship date");
                WriteUnicodeString(stream, "Invalid date");
                WriteUnicodeString(stream, "Enter a 2024 ship date.");
                WriteUnicodeString(stream, "Use a date in 2024.");
                WriteDvFormula(stream, BuildPtgNumFormula(new DateTime(2024, 1, 1).ToOADate()));
                WriteDvFormula(stream, BuildPtgNumFormula(new DateTime(2024, 12, 31).ToOADate()));
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 4);
                return stream.ToArray();
            }

            private static byte[] BuildTimeDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x05U
                    | (0x06U << 20)
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, string.Empty);
                WriteUnicodeString(stream, "Invalid time");
                WriteUnicodeString(stream, string.Empty);
                WriteUnicodeString(stream, "Use a time at or after 09:00.");
                WriteDvFormula(stream, BuildPtgNumFormula(TimeSpan.FromHours(9).TotalDays));
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 5);
                WriteUInt16(stream, 5);
                return stream.ToArray();
            }

            private static byte[] BuildTextLengthDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x06U
                    | (0x07U << 20)
                    | 0x00000100U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, string.Empty);
                WriteUnicodeString(stream, "Invalid text");
                WriteUnicodeString(stream, string.Empty);
                WriteUnicodeString(stream, "Use 12 characters or fewer.");
                WriteDvFormula(stream, BuildPtgIntFormula(12));
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 6);
                WriteUInt16(stream, 6);
                return stream.ToArray();
            }

            private static byte[] BuildCustomFormulaDataValidationPayload() {
                using var stream = new MemoryStream();
                uint flags = 0x07U
                    | 0x00000100U
                    | 0x00040000U
                    | 0x00080000U;
                WriteUInt32(stream, flags);
                WriteUnicodeString(stream, "Custom");
                WriteUnicodeString(stream, "Invalid total");
                WriteUnicodeString(stream, "Enter a value allowed by the custom formula.");
                WriteUnicodeString(stream, "The custom formula rejected this value.");
                WriteDvFormula(stream, BuildCustomValidationFormula());
                WriteDvFormula(stream, Array.Empty<byte>());
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 10);
                WriteUInt16(stream, 5);
                WriteUInt16(stream, 5);
                return stream.ToArray();
            }

            private static byte[] BuildCustomValidationFormula() {
                using var stream = new MemoryStream();
                byte[] area = BuildPtgAreaFormula(0, 0, 0, 1);
                stream.Write(area, 0, area.Length);
                stream.WriteByte(0x42);
                stream.WriteByte(0x01);
                WriteUInt16(stream, 0x0004);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, 10);
                stream.WriteByte(0x0d);
                return stream.ToArray();
            }

            private static byte[] BuildPtgIntFormula(ushort value) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x1e);
                WriteUInt16(stream, value);
                return stream.ToArray();
            }

            private static byte[] BuildPtgNumFormula(double value) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x1f);
                byte[] valueBytes = BitConverter.GetBytes(value);
                stream.Write(valueBytes, 0, valueBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildPtgStrFormula(string value) {
                byte[] valueBytes = Encoding.ASCII.GetBytes(value);
                using var stream = new MemoryStream();
                stream.WriteByte(0x17);
                stream.WriteByte(checked((byte)value.Length));
                stream.WriteByte(0);
                stream.Write(valueBytes, 0, valueBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildPtgAreaFormula(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x25);
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
                return stream.ToArray();
            }

            private static byte[] BuildPtgArea3dFormula(ushort externSheetIndex, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x3b);
                WriteUInt16(stream, externSheetIndex);
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
                return stream.ToArray();
            }

            private static byte[] BuildPtgNameFormula(uint oneBasedNameIndex) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x23);
                WriteUInt32(stream, oneBasedNameIndex);
                return stream.ToArray();
            }

            private static void WriteDvFormula(Stream stream, byte[] formulaBytes) {
                WriteUInt16(stream, checked((ushort)formulaBytes.Length));
                WriteUInt16(stream, 0);
                stream.Write(formulaBytes, 0, formulaBytes.Length);
            }

            private static void WriteUnicodeString(Stream stream, string text) {
                WriteUInt16(stream, checked((ushort)text.Length));
                stream.WriteByte(0x01);
                byte[] textBytes = Encoding.Unicode.GetBytes(text);
                stream.Write(textBytes, 0, textBytes.Length);
            }

            private static void WriteCompressedUnicodeStringNoCch(Stream stream, string value) {
                stream.WriteByte(0);
                byte[] valueBytes = Encoding.ASCII.GetBytes(value);
                stream.Write(valueBytes, 0, valueBytes.Length);
            }

            private static byte[] BuildDoublePayload(double value) {
                return BitConverter.GetBytes(value);
            }

            private static byte[] BuildGutsPayload(ushort rowOutlineLevelRaw, ushort columnOutlineLevelRaw) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, rowOutlineLevelRaw);
                WriteUInt16(stream, columnOutlineLevelRaw);
                return stream.ToArray();
            }

            private static byte[] BuildIndexPayload(uint firstRow, uint rowAfterLast, uint reservedRecordOffset, uint[] dbCellOffsets) {
                using var stream = new MemoryStream();
                WriteUInt32(stream, 0);
                WriteUInt32(stream, firstRow);
                WriteUInt32(stream, rowAfterLast);
                WriteUInt32(stream, reservedRecordOffset);
                foreach (uint offset in dbCellOffsets) {
                    WriteUInt32(stream, offset);
                }

                return stream.ToArray();
            }

            private static byte[] BuildSelectionPayload() {
                using var stream = new MemoryStream();
                stream.WriteByte(0);
                WriteUInt16(stream, 2);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 2);
                WriteUInt16(stream, 3);
                stream.WriteByte(1);
                stream.WriteByte(2);
                return stream.ToArray();
            }

            private static byte[] BuildSortPayload(string key1, string key2, string key3) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x047b);
                stream.WriteByte(checked((byte)key1.Length));
                stream.WriteByte(checked((byte)key2.Length));
                stream.WriteByte(checked((byte)key3.Length));
                WriteCompressedUnicodeStringNoCch(stream, key1);
                WriteCompressedUnicodeStringNoCch(stream, key2);
                WriteCompressedUnicodeStringNoCch(stream, key3);
                stream.WriteByte(0);
                return stream.ToArray();
            }

            private static byte[] BuildSxdiPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 2);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 7);
                WriteUInt16(stream, 0xffff);
                WriteUInt16(stream, 0xffff);
                WriteUInt16(stream, 14);
                WriteUInt16(stream, 5);
                WriteCompressedUnicodeStringNoCch(stream, "Sales");
                return stream.ToArray();
            }

            private static byte[] BuildSxRngPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0017);
                return stream.ToArray();
            }

            private static byte[] BuildSxVdExPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x009f);
                WriteUInt16(stream, 10);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 14);
                return stream.ToArray();
            }

            private static byte[] BuildSetupPayload(
                ushort scale,
                ushort fitToWidth,
                ushort fitToHeight,
                bool landscape,
                double header,
                double footer) {
                byte[] payload = new byte[34];
                WriteUInt16(payload, 0, 1);
                WriteUInt16(payload, 2, scale);
                WriteUInt16(payload, 6, fitToWidth);
                WriteUInt16(payload, 8, fitToHeight);
                WriteUInt16(payload, 10, landscape ? (ushort)0x0000 : (ushort)0x0002);
                Buffer.BlockCopy(BitConverter.GetBytes(header), 0, payload, 16, 8);
                Buffer.BlockCopy(BitConverter.GetBytes(footer), 0, payload, 24, 8);
                WriteUInt16(payload, 32, 1);
                return payload;
            }

            private static byte[] BuildUInt16Payload(ushort value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, value);
                return stream.ToArray();
            }

            private static byte[] BuildUnicodeStringPayload(string text) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)text.Length));
                stream.WriteByte(0x01);
                byte[] textBytes = Encoding.Unicode.GetBytes(text);
                stream.Write(textBytes, 0, textBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildExternSheetPayload(params (ushort SupBookIndex, short FirstSheetIndex, short LastSheetIndex)[] references) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)references.Length));
                foreach ((ushort supBookIndex, short firstSheetIndex, short lastSheetIndex) in references) {
                    WriteUInt16(stream, supBookIndex);
                    WriteUInt16(stream, unchecked((ushort)firstSheetIndex));
                    WriteUInt16(stream, unchecked((ushort)lastSheetIndex));
                }

                return stream.ToArray();
            }

            private static byte[] BuildDefinedNamePayload(string name, byte[] formula, ushort localSheetIndex, bool hidden, bool builtIn) {
                byte[] nameBytes = Encoding.ASCII.GetBytes(name);
                using var stream = new MemoryStream();
                ushort flags = 0;
                if (hidden) {
                    flags |= 0x0001;
                }

                if (builtIn) {
                    flags |= 0x0020;
                }

                WriteUInt16(stream, flags);
                stream.WriteByte(0);
                stream.WriteByte(checked((byte)name.Length));
                WriteUInt16(stream, checked((ushort)formula.Length));
                WriteUInt16(stream, 0);
                WriteUInt16(stream, localSheetIndex);
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.Write(nameBytes, 0, nameBytes.Length);
                stream.Write(formula, 0, formula.Length);
                return stream.ToArray();
            }

            private static byte[] BuildNameRef3dFormula(ushort externSheetIndex, ushort row, ushort column) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x3a);
                WriteUInt16(stream, externSheetIndex);
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                return stream.ToArray();
            }

            private static byte[] BuildNameArea3dFormula(ushort externSheetIndex, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x3b);
                WriteUInt16(stream, externSheetIndex);
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
                return stream.ToArray();
            }

            private static byte[] BuildNamePrintTitlesFormula(ushort externSheetIndex) {
                using var stream = new MemoryStream();
                byte[] titleRows = BuildNameArea3dFormula(externSheetIndex, 0, 0, 1, 255);
                byte[] titleColumns = BuildNameArea3dFormula(externSheetIndex, 0, 0, ushort.MaxValue, 1);
                stream.Write(titleRows, 0, titleRows.Length);
                stream.Write(titleColumns, 0, titleColumns.Length);
                stream.WriteByte(0x10);
                return stream.ToArray();
            }
        }
    }
}
