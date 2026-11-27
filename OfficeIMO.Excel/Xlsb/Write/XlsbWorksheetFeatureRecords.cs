using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>What a worksheet feature writer can use while it encodes one generated worksheet.</summary>
    internal sealed class XlsbWorksheetFeatureContext {
        internal XlsbWorksheetFeatureContext(
            ExcelDocument document,
            ExcelSheet sheet,
            IXlsbFormulaEncodingContext? formulas,
            XlsbWorksheetPackageParts? parts,
            XlsbRichTextFontTable? richText = null) {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            Formulas = formulas;
            Parts = parts;
            RichText = richText;
        }

        internal ExcelDocument Document { get; }

        internal ExcelSheet Sheet { get; }

        internal Worksheet Worksheet => Sheet.WorksheetPart.Worksheet
            ?? throw new InvalidDataException($"Worksheet '{Sheet.Name}' has no worksheet root.");

        internal OfficeXmlNode WorksheetMarkup => XlsbWorksheetMarkup.Root(Sheet);

        /// <summary>Encodes formula text against the workbook's sheets and names; null where formulas cannot be encoded.</summary>
        internal IXlsbFormulaEncodingContext? Formulas { get; }

        /// <summary>Collects related parts and worksheet relationships; null where the writer cannot add parts.</summary>
        internal XlsbWorksheetPackageParts? Parts { get; }

        /// <summary>Maps formatted text runs to BrtFont indexes, adding fonts where needed; null where fonts cannot be added.</summary>
        internal XlsbRichTextFontTable? RichText { get; }

        /// <summary>
        /// Encodes <paramref name="formulaText"/> and returns the shared cce/rgce/cb/rgcb layout of CellParsedFormula
        /// without the leading grbitFlags, the layout DVParsedFormula, CFParsedFormula and similar structures use.
        /// </summary>
        /// <remarks>resultClass is the operand class of the result (ExcelNativeFormulaOperandTypes); value class by default.</remarks>
        internal byte[] EncodeFormula(string formulaText, string purpose, byte resultClass = ExcelNativeFormulaOperandTypes.Value) {
            if (Formulas == null) {
                throw new NotSupportedException($"Native XLSB generation cannot encode the {purpose} formula on worksheet '{Sheet.Name}'.");
            }
            string text = formulaText.StartsWith("=", StringComparison.Ordinal) ? formulaText.Substring(1) : formulaText;
            if (!Formulas.TryEncode(text, Sheet.Name, out byte[] payload, out string? reason, resultClass: resultClass)) {
                throw new NotSupportedException($"Native XLSB generation cannot encode the {purpose} formula '{formulaText}' on worksheet '{Sheet.Name}': {reason}");
            }
            if (payload.Length < 2) throw new InvalidDataException("The native formula encoder returned a truncated payload.");
            var result = new byte[payload.Length - 2];
            Buffer.BlockCopy(payload, 2, result, 0, result.Length);
            return result;
        }
    }

    /// <summary>
    /// Generated BIFF12 records for worksheet content written after the cell table, grouped by their position in the
    /// MS-XLSB WORKSHEET grammar (section 2.1.7.62). <see cref="XlsbWorksheetPartWriter"/> writes each group in place.
    /// </summary>
    internal sealed class XlsbWorksheetFeatureRecords {
        internal static readonly XlsbWorksheetFeatureRecords Empty = new XlsbWorksheetFeatureRecords();

        /// <summary>*CONDITIONALFORMATTING, after MERGECELLS and BrtPhoneticInfo.</summary>
        internal List<XlsbGeneratedRecord> ConditionalFormatting { get; } = new List<XlsbGeneratedRecord>();

        /// <summary>[DVALS], before the hyperlinks.</summary>
        internal List<XlsbGeneratedRecord> DataValidations { get; } = new List<XlsbGeneratedRecord>();

        /// <summary>[RWBRK] [COLBRK], after HEADERFOOTER.</summary>
        internal List<XlsbGeneratedRecord> PageBreaks { get; } = new List<XlsbGeneratedRecord>();

        /// <summary>[IGNOREECS], after CELLWATCHES.</summary>
        internal List<XlsbGeneratedRecord> IgnoredErrors { get; } = new List<XlsbGeneratedRecord>();

        /// <summary>[BrtDrawing] [BrtLegacyDrawing] [BrtLegacyDrawingHF], after SMARTTAGS.</summary>
        internal List<XlsbGeneratedRecord> Drawings { get; } = new List<XlsbGeneratedRecord>();

        /// <summary>[LISTPARTS], after WEBPUBITEMS.</summary>
        internal List<XlsbGeneratedRecord> TableParts { get; } = new List<XlsbGeneratedRecord>();

        /// <summary>FRTWORKSHEET future records, immediately before BrtEndSheet.</summary>
        internal List<XlsbGeneratedRecord> FutureRecords { get; } = new List<XlsbGeneratedRecord>();

        internal static XlsbWorksheetFeatureRecords Create(XlsbWorksheetFeatureContext context) {
            ArgumentNullException.ThrowIfNull(context);
            var records = new XlsbWorksheetFeatureRecords();
            XlsbWorksheetConditionalFormattingWriter.Append(records, context);
            XlsbWorksheetPageBreakWriter.Append(records, context);
            XlsbWorksheetIgnoredErrorWriter.Append(records, context);
            XlsbWorksheetDataValidationWriter.Append(records, context);
            // BrtDrawing precedes the comments' BrtLegacyDrawing in the WORKSHEET grammar.
            XlsbWorksheetDrawingWriter.Append(records, context);
            XlsbWorksheetCommentWriter.Append(records, context);
            XlsbWorksheetTableWriter.Append(records, context);
            return records;
        }
    }
}
