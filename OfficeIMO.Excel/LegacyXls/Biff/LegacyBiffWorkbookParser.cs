using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class LegacyBiffWorkbookParser {
        internal static LegacyXlsWorkbook Parse(byte[] workbookStream, LegacyXlsImportOptions options) {
            var workbook = new LegacyXlsWorkbook();
            IReadOnlyList<BiffRecord> records = ReadWorkbookGlobalRecords(workbookStream, workbook.MutableDiagnostics);
            var sharedStrings = new List<string>();
            var numberFormatsById = new Dictionary<ushort, string>();
            var externSheets = new List<BiffExternSheetReference>();
            var boundSheetNames = new List<string>();
            var boundSheetProjectedWorksheetIndexes = new List<int?>();
            var definedNameTable = new List<string?>();
            LegacyXlsExternalReference? currentExternalReference = null;
            LegacyXlsExternalCellCache? currentExternalCellCache = null;
            var chartMetadataState = new BiffChartMetadataReaderState();
            var pivotTableMetadataState = new BiffPivotTableMetadataReaderState();

            if (!LegacyBiffVersionValidator.ValidateWorkbookGlobals(records, workbook)) {
                return workbook;
            }

            for (int i = 0; i < records.Count; i++) {
                BiffRecord record = records[i];
                if (record.Type == (ushort)BiffRecordType.BoundSheet8) {
                    LegacyXlsWorksheet? sheet = TryReadBoundSheet(record, workbook.MutableDiagnostics);
                    if (sheet != null) {
                        boundSheetNames.Add(sheet.Name);
                    }

                    if (sheet != null && sheet.SheetType == 0) {
                        boundSheetProjectedWorksheetIndexes.Add(workbook.MutableWorksheets.Count);
                        workbook.MutableWorksheets.Add(sheet);
                    } else if (sheet != null) {
                        boundSheetProjectedWorksheetIndexes.Add(null);
                        LegacyXlsUnsupportedSheet unsupportedSheet = ToUnsupportedSheet(sheet, ToUnsupportedSheetKind(sheet.SheetType));
                        workbook.MutableUnsupportedSheets.Add(unsupportedSheet);
                        workbook.MutableUnsupportedFeatures.Add(BiffUnsupportedRecordDiagnostics.CreateUnsupportedSheetTypeFeature(record, unsupportedSheet));
                        if (options.ReportUnsupportedRecords) {
                            BiffUnsupportedRecordDiagnostics.AddUnsupportedSheetTypeDiagnostic(workbook.MutableDiagnostics, record, unsupportedSheet);
                        }
                    }
                } else if (record.Type == (ushort)BiffRecordType.Date1904) {
                    ReadDate1904(record, workbook, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.ExternSheet) {
                    ReadExternSheet(record, externSheets, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.FilePass) {
                    workbook.MutableUnsupportedFeatures.Add(BiffUnsupportedRecordDiagnostics.CreateFilePassFeature(record));
                    BiffUnsupportedRecordDiagnostics.AddFilePassDiagnostic(record, workbook.MutableDiagnostics);
                    return workbook;
                } else if (record.Type == (ushort)BiffRecordType.Font) {
                    ReadFont(record, workbook, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.Format) {
                    ReadFormat(record, workbook, numberFormatsById, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.Lbl) {
                    ReadDefinedName(record, workbook, externSheets, workbook.ExternalReferences, boundSheetNames, boundSheetProjectedWorksheetIndexes, definedNameTable, workbook.MutableFormulaTokenRecords, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.Palette) {
                    ReadPalette(record, workbook, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.Password) {
                    ReadPassword(record, workbook);
                } else if (record.Type == (ushort)BiffRecordType.Protect) {
                    ReadProtect(record, workbook);
                } else if (record.Type == (ushort)BiffRecordType.Sst) {
                    IReadOnlyList<byte[]> payloads = CollectContinuedPayloads(records, ref i);
                    sharedStrings = BiffStringReader.ReadSharedStrings(payloads, workbook.MutableDiagnostics, record.Offset);
                } else if (record.Type == (ushort)BiffRecordType.SupBook) {
                    currentExternalReference = ReadSupBook(record, workbook, workbook.MutableDiagnostics, options);
                    currentExternalCellCache = null;
                } else if (record.Type == (ushort)BiffRecordType.ExternName) {
                    ReadExternName(record, currentExternalReference, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.Xct) {
                    currentExternalCellCache = BiffExternalCellCacheReader.ReadXct(record, currentExternalReference, workbook.MutableDiagnostics);
                } else if (record.Type == (ushort)BiffRecordType.Crn) {
                    BiffExternalCellCacheReader.ReadCrn(record, currentExternalCellCache, workbook.MutableDiagnostics);
                } else if (BiffCalculationSettingsReader.TryRead(record, sheetName: null, workbook.MutableCalculationSettings, workbook.MutableDiagnostics)) {
                    continue;
                } else if (BiffWorkbookMetadataReader.TryRead(record, workbook, workbook.MutableDiagnostics)) {
                    continue;
                } else if (BiffThemeReader.TryRead(record, workbook.MutableDiagnostics, out LegacyXlsThemeRecord? themeRecord)) {
                    if (themeRecord != null) {
                        workbook.MutableThemeRecords.Add(themeRecord);
                    }

                    AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                    }

                    continue;
                } else if (BiffStyleReader.TryRead(record, workbook, workbook.MutableDiagnostics)) {
                    if (record.Type == (ushort)BiffRecordType.XfCrc
                        || record.Type == (ushort)BiffRecordType.XfExt
                        || record.Type == (ushort)BiffRecordType.StyleExt) {
                        AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                        if (options.ReportUnsupportedRecords) {
                            BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                        }
                    }

                    continue;
                } else if (BiffDrawingMetadataReader.TryRead(record, sheetName: null, workbook.MutableDrawingRecords)) {
                    AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                    }
                } else if (BiffDataConsolidationReferenceReader.TryRead(record, workbook.MutableDiagnostics, out LegacyXlsDataConsolidationReference? dataConsolidationReference)) {
                    workbook.MutableDataConsolidationReferences.Add(dataConsolidationReference!);
                    AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                    }
                } else if (BiffChartMetadataReader.TryRead(record, sheetName: null, workbook.MutableChartRecords, chartMetadataState)) {
                    BiffChartMetadataReader.ScanFormulaTokens(record, sheetName: null, workbook.MutableFormulaTokenRecords);
                    AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                    }
                } else if (BiffPivotTableMetadataReader.TryRead(record, sheetName: null, workbook.MutablePivotTableRecords, workbook.MutableDiagnostics, pivotTableMetadataState)) {
                    AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                    }
                } else if (record.Type == (ushort)BiffRecordType.Dxf) {
                    if (BiffDifferentialFormatReader.TryRead(record, workbook, out LegacyXlsDifferentialFormat? differentialFormat)) {
                        workbook.MutableDifferentialFormats.Add(differentialFormat!);
                    }

                    AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                    }
                } else if (record.Type == (ushort)BiffRecordType.Xf) {
                    ReadCellFormat(record, workbook, numberFormatsById, workbook.MutableDiagnostics);
                } else if (record.Type != (ushort)BiffRecordType.Bof && record.Type != (ushort)BiffRecordType.Eof) {
                    AddUnsupportedRecordFeature(workbook, record, sheetName: null);
                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(workbook.MutableDiagnostics, record.Type, record.Offset, sheetName: null);
                    }
                }
            }

            MoveDialogSheetsToUnsupported(workbookStream, workbook, options);
            LegacyBiffUnsupportedSheetScanner.Scan(
                workbookStream,
                workbook.UnsupportedSheets,
                workbook.MutableUnsupportedFeatures,
                workbook.MutablePreservedFeatureRecords,
                workbook.MutablePivotTableRecords,
                workbook.MutableChartRecords,
                workbook.MutableDrawingRecords,
                workbook.MutableFormulaTokenRecords,
                workbook.MutableDiagnostics,
                options);

            IReadOnlyList<string> sheetNames = boundSheetNames.Count == 0
                ? workbook.Worksheets.Select(sheet => sheet.Name).ToArray()
                : boundSheetNames.ToArray();
            foreach (LegacyXlsWorksheet sheet in workbook.Worksheets) {
                LegacyBiffWorksheetParser.Parse(workbookStream, sheet, sharedStrings, externSheets, workbook.ExternalReferences, sheetNames, definedNameTable, workbook.MutableUnsupportedFeatures, workbook.MutablePreservedFeatureRecords, workbook.MutablePivotTableRecords, workbook.MutableChartRecords, workbook.MutableDrawingRecords, workbook.DifferentialFormats, workbook.MutableCalculationSettings, workbook.MutableFormulaTokenRecords, workbook.MutableDiagnostics, options);
            }

            return workbook;
        }

        private static void AddUnsupportedRecordFeature(
            LegacyXlsWorkbook workbook,
            BiffRecord record,
            string? sheetName) {
            LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(record.Type, record.Offset, sheetName);
            workbook.MutableUnsupportedFeatures.Add(feature);
            if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, record.Payload.Length, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                workbook.MutablePreservedFeatureRecords.Add(preservedRecord!);
            }
        }

        private static IReadOnlyList<byte[]> CollectContinuedPayloads(IReadOnlyList<BiffRecord> records, ref int index) {
            BiffRecord first = records[index];
            int lastIndex = index;
            while (lastIndex + 1 < records.Count && records[lastIndex + 1].Type == (ushort)BiffRecordType.Continue) {
                lastIndex++;
            }

            if (lastIndex == index) {
                return new[] { first.Payload };
            }

            var payloads = new byte[checked(lastIndex - index + 1)][];
            for (int i = index; i <= lastIndex; i++) {
                payloads[i - index] = records[i].Payload;
            }

            index = lastIndex;
            return payloads;
        }

        private static IReadOnlyList<BiffRecord> ReadWorkbookGlobalRecords(byte[] workbookStream, List<LegacyXlsImportDiagnostic> diagnostics) {
            var records = new List<BiffRecord>();
            int offset = 0;
            while (offset < workbookStream.Length) {
                if (offset + 4 > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Error,
                        "XLS-BIFF-GLOBALS-TRUNCATED-HEADER",
                        "The BIFF globals stream ended inside a record header.",
                        recordOffset: offset));
                    break;
                }

                ushort type = BiffRecordReader.ReadUInt16(workbookStream, offset);
                ushort length = BiffRecordReader.ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Error,
                        "XLS-BIFF-GLOBALS-TRUNCATED-PAYLOAD",
                        $"BIFF record 0x{type:X4} declares {length} payload bytes, but the stream ends early.",
                        recordOffset: offset,
                        recordType: type));
                    break;
                }

                byte[] payload = new byte[length];
                Buffer.BlockCopy(workbookStream, payloadOffset, payload, 0, length);
                records.Add(new BiffRecord(type, offset, payload));
                offset = payloadOffset + length;
                if (type == (ushort)BiffRecordType.Eof) {
                    break;
                }
            }

            return records;
        }

        private static void ReadDate1904(BiffRecord record, LegacyXlsWorkbook workbook, List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-DATE1904-SHORT",
                    "A Date1904 record was too short to parse.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            workbook.SetUses1904DateSystem(BiffRecordReader.ReadUInt16(record.Payload, 0) != 0);
        }

        private static void ReadPassword(BiffRecord record, LegacyXlsWorkbook workbook) {
            if (record.Payload.Length < 2) {
                return;
            }

            ushort passwordHash = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (passwordHash != 0) {
                workbook.SetProtectionPasswordHash(passwordHash);
            }
        }

        private static void ReadProtect(BiffRecord record, LegacyXlsWorkbook workbook) {
            if (record.Payload.Length < 2) {
                return;
            }

            workbook.SetProtection(BiffRecordReader.ReadUInt16(record.Payload, 0) != 0);
        }

        private static LegacyXlsExternalReference? ReadSupBook(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options) {
            if (!BiffSupBookReader.TryRead(record, diagnostics, out LegacyXlsExternalReference? reference)) {
                return null;
            }

            workbook.MutableExternalReferences.Add(reference!);
            bool unsupportedExternalReference = reference!.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook
                || reference.Kind == LegacyXlsExternalReferenceKind.AddIn
                || reference.Kind == LegacyXlsExternalReferenceKind.DdeOrOle;
            if (unsupportedExternalReference) {
                LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateExternalReferenceFeature(record, reference);
                workbook.MutableUnsupportedFeatures.Add(feature);
                if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, record.Payload.Length, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                    workbook.MutablePreservedFeatureRecords.Add(preservedRecord!);
                }

                if (options.ReportUnsupportedRecords) {
                    BiffUnsupportedRecordDiagnostics.AddExternalReferenceDiagnostic(diagnostics, record, reference);
                }
            }

            return reference;
        }

        private static void ReadExternName(
            BiffRecord record,
            LegacyXlsExternalReference? currentExternalReference,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (currentExternalReference == null) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-EXTERNNAME-ORPHANED",
                    "An ExternName record appeared before a SupBook supporting link.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            if (BiffExternNameReader.TryRead(record, currentExternalReference.Kind, diagnostics, out LegacyXlsExternalName? externalName)) {
                currentExternalReference.MutableExternalNames.Add(externalName!);
            }
        }

        private static void ReadFormat(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            Dictionary<ushort, string> numberFormatsById,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            try {
                if (record.Payload.Length < 5) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-FORMAT-SHORT",
                        "A Format record was too short to parse.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return;
                }

                ushort formatId = BiffRecordReader.ReadUInt16(record.Payload, 0);
                int stringOffset = 2;
                string formatCode = BiffStringReader.ReadUnicodeString(record.Payload, ref stringOffset);
                numberFormatsById[formatId] = formatCode;
                workbook.MutableNumberFormats.Add(new LegacyXlsNumberFormat(formatId, formatCode));
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-FORMAT-INVALID",
                    $"A Format record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }
        }

        private static void ReadExternSheet(
            BiffRecord record,
            List<BiffExternSheetReference> externSheets,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-EXTERNSHEET-SHORT",
                    "An ExternSheet record was too short to parse.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            ushort count = BiffRecordReader.ReadUInt16(record.Payload, 0);
            int expectedLength = checked(2 + (count * 6));
            if (expectedLength > record.Payload.Length) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-EXTERNSHEET-INVALID",
                    "An ExternSheet record ended before all XTI references could be read.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            externSheets.Clear();
            for (int i = 0; i < count; i++) {
                int offset = 2 + (i * 6);
                externSheets.Add(new BiffExternSheetReference(
                    BiffRecordReader.ReadUInt16(record.Payload, offset),
                    unchecked((short)BiffRecordReader.ReadUInt16(record.Payload, offset + 2)),
                    unchecked((short)BiffRecordReader.ReadUInt16(record.Payload, offset + 4))));
            }
        }

        private static void ReadDefinedName(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<int?> boundSheetProjectedWorksheetIndexes,
            List<string?> definedNameTable,
            List<LegacyXlsFormulaTokenRecord> formulaTokenRecords,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            string? formulaName = null;
            try {
                if (record.Payload.Length < 14) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-LBL-SHORT",
                        "A Lbl record was too short to parse.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return;
                }

                ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 0);
                bool hidden = (flags & 0x0001) != 0;
                bool builtIn = (flags & 0x0020) != 0;
                int nameCharCount = record.Payload[3];
                ushort formulaLength = BiffRecordReader.ReadUInt16(record.Payload, 4);
                ushort oneBasedSheetIndex = BiffRecordReader.ReadUInt16(record.Payload, 8);
                int offset = 14;
                string rawName = BiffStringReader.ReadUnicodeStringNoCch(record.Payload, ref offset, nameCharCount);
                if (offset + formulaLength > record.Payload.Length) {
                    throw new InvalidDataException("The Lbl record ended before the parsed formula could be read.");
                }

                string? name = builtIn ? GetBuiltInName(rawName) : rawName;
                formulaName = name;
                if (string.IsNullOrWhiteSpace(name)) {
                    return;
                }

                byte[] formulaBytes = new byte[formulaLength];
                Buffer.BlockCopy(record.Payload, offset, formulaBytes, 0, formulaLength);
                BiffFormulaTokenScanner.ScanTokens(
                    formulaBytes,
                    "DefinedName",
                    sheetName: null,
                    cellReference: name,
                    record.Offset,
                    record.Type,
                    formulaTokenRecords);
                string? reference;
                BiffFormulaReadFailure? formulaFailure;
                bool formulaRead = string.Equals(name, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase)
                    ? BiffNameFormulaReader.TryReadPrintTitles(formulaBytes, externSheets, externalReferences, sheetNames, out reference, out formulaFailure)
                    : BiffNameFormulaReader.TryReadReference(formulaBytes, externSheets, externalReferences, sheetNames, out reference, out formulaFailure);
                if (!formulaRead) {
                    if (TryAddFutureFunctionAlias(workbook, name!, record, formulaFailure)) {
                        return;
                    }

                    if (formulaFailure == null) {
                        diagnostics.Add(new LegacyXlsImportDiagnostic(
                            LegacyXlsDiagnosticSeverity.Info,
                            "XLS-BIFF-LBL-FORMULA-UNSUPPORTED",
                            $"Defined name '{name}' uses a formula shape that is not imported yet.",
                            recordOffset: record.Offset,
                            recordType: record.Type));
                    } else {
                        diagnostics.Add(new LegacyXlsImportDiagnostic(
                            LegacyXlsDiagnosticSeverity.Info,
                            "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED",
                            $"{formulaFailure.Description} Defined name '{name}' uses a formula shape that is not imported yet.",
                            recordOffset: record.Offset,
                            recordType: record.Type,
                            detailCode: formulaFailure.DetailCode,
                            formulaContext: "DefinedName",
                            formulaToken: formulaFailure.Token,
                            formulaTokenName: formulaFailure.TokenName,
                            formulaTokenOffset: formulaFailure.TokenOffset));
                    }

                    return;
                }

                int? localSheetIndex = null;
                if (oneBasedSheetIndex != 0) {
                    int boundSheetIndex = oneBasedSheetIndex - 1;
                    if (boundSheetIndex < 0 || boundSheetIndex >= boundSheetProjectedWorksheetIndexes.Count) {
                        diagnostics.Add(new LegacyXlsImportDiagnostic(
                            LegacyXlsDiagnosticSeverity.Warning,
                            "XLS-BIFF-LBL-SCOPE-INVALID",
                            $"Defined name '{name}' references a sheet scope outside the workbook.",
                            recordOffset: record.Offset,
                            recordType: record.Type));
                        return;
                    }

                    localSheetIndex = boundSheetProjectedWorksheetIndexes[boundSheetIndex];
                    if (!localSheetIndex.HasValue) {
                        diagnostics.Add(new LegacyXlsImportDiagnostic(
                            LegacyXlsDiagnosticSeverity.Info,
                            "XLS-BIFF-LBL-SCOPE-UNSUPPORTED",
                            $"Defined name '{name}' is scoped to a sheet type that is not imported yet.",
                            recordOffset: record.Offset,
                            recordType: record.Type));
                        return;
                    }
                }

                if (localSheetIndex.HasValue && (localSheetIndex.Value < 0 || localSheetIndex.Value >= workbook.Worksheets.Count)) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-LBL-SCOPE-INVALID",
                        $"Defined name '{name}' references a sheet scope outside the workbook.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return;
                }

                workbook.MutableDefinedNames.Add(new LegacyXlsDefinedName(name!, reference!, localSheetIndex, hidden, builtIn));
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-LBL-INVALID",
                    $"A Lbl record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            } finally {
                definedNameTable.Add(formulaName);
            }
        }

        private static bool TryAddFutureFunctionAlias(LegacyXlsWorkbook workbook, string name, BiffRecord record, BiffFormulaReadFailure? formulaFailure) {
            const string futureFunctionPrefix = "_xlfn.";
            if (!name.StartsWith(futureFunctionPrefix, StringComparison.OrdinalIgnoreCase)
                || formulaFailure?.Token != 0x1c) {
                return false;
            }

            workbook.MutableFutureFunctionAliases.Add(new LegacyXlsFutureFunctionAlias(
                name,
                name.Substring(futureFunctionPrefix.Length),
                record.Offset,
                record.Type,
                formulaFailure.Token,
                formulaFailure.TokenName,
                formulaFailure.TokenOffset));
            return true;
        }

        private static string? GetBuiltInName(string rawName) {
            if (rawName.Length != 1) {
                return null;
            }

            switch (rawName[0]) {
                case (char)0x06:
                    return "_xlnm.Print_Area";
                case (char)0x07:
                    return "_xlnm.Print_Titles";
                case (char)0x0d:
                    return "_FilterDatabase";
                default:
                    return null;
            }
        }

        private static void ReadFont(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            try {
                if (record.Payload.Length < 16) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-FONT-SHORT",
                        "A Font record was too short to parse.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return;
                }

                ushort fontIndex = checked((ushort)workbook.MutableFonts.Count);
                ushort heightTwips = BiffRecordReader.ReadUInt16(record.Payload, 0);
                ushort options = BiffRecordReader.ReadUInt16(record.Payload, 2);
                ushort colorIndex = BiffRecordReader.ReadUInt16(record.Payload, 4);
                ushort weight = BiffRecordReader.ReadUInt16(record.Payload, 6);
                ushort escapement = BiffRecordReader.ReadUInt16(record.Payload, 8);
                byte underline = record.Payload[10];
                int nameOffset = 14;
                string name = BiffStringReader.ReadShortUnicodeString(record.Payload, ref nameOffset);

                workbook.MutableFonts.Add(new LegacyXlsFont(
                    fontIndex,
                    string.IsNullOrWhiteSpace(name) ? null : name,
                    heightTwips == 0 ? null : heightTwips / 20d,
                    colorIndex,
                    weight >= 700,
                    (options & 0x0002) != 0,
                    underline != 0,
                    (options & 0x0008) != 0,
                    ToFontEscapement(escapement)));
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-FONT-INVALID",
                    $"A Font record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }
        }

        private static LegacyXlsFontEscapement ToFontEscapement(ushort escapement) {
            return escapement == 1
                ? LegacyXlsFontEscapement.Superscript
                : escapement == 2
                    ? LegacyXlsFontEscapement.Subscript
                    : LegacyXlsFontEscapement.None;
        }

        private static void ReadPalette(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-PALETTE-SHORT",
                    "A Palette record was too short to parse.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            ushort colorCount = BiffRecordReader.ReadUInt16(record.Payload, 0);
            int expectedLength = checked(2 + (colorCount * 4));
            if (expectedLength > record.Payload.Length) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-PALETTE-INVALID",
                    "A Palette record ended before all colors could be read.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            workbook.MutablePaletteColors.Clear();
            for (int i = 0; i < colorCount; i++) {
                int offset = 2 + (i * 4);
                workbook.MutablePaletteColors.Add(
                    "FF"
                    + record.Payload[offset].ToString("X2", System.Globalization.CultureInfo.InvariantCulture)
                    + record.Payload[offset + 1].ToString("X2", System.Globalization.CultureInfo.InvariantCulture)
                    + record.Payload[offset + 2].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        private static void ReadCellFormat(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            IReadOnlyDictionary<ushort, string> numberFormatsById,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < 4) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XF-SHORT",
                    "An XF record was too short to parse.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            ushort styleIndex = checked((ushort)workbook.MutableCellFormats.Count);
            ushort fontIndex = BiffRecordReader.ReadUInt16(record.Payload, 0);
            ushort numberFormatId = BiffRecordReader.ReadUInt16(record.Payload, 2);
            ushort protection = BiffRecordReader.ReadUInt16(record.Payload, 4);
            bool isStyle = (protection & 0x0004) != 0;
            ushort parentStyleIndex = (ushort)((protection >> 4) & 0x0fff);
            ReadCellFormatApplyFlags(
                record.Payload,
                out bool applyNumberFormat,
                out bool applyFont);
            ReadCellFormatFill(record.Payload, out byte fillPattern, out ushort fillForegroundColorIndex, out ushort fillBackgroundColorIndex);
            ReadCellFormatAlignment(
                record.Payload,
                out bool applyAlignment,
                out byte horizontalAlignment,
                out byte verticalAlignment,
                out bool wrapText,
                out byte textRotation,
                out byte indent,
                out bool shrinkToFit,
                out byte readingOrder);
            ReadCellFormatProtection(
                record.Payload,
                out bool applyProtection,
                out bool locked,
                out bool formulaHidden,
                out bool quotePrefix);
            LegacyXlsBorder? border = ReadCellFormatBorder(record.Payload);
            bool isBuiltIn = BiffBuiltInNumberFormat.TryGetCode(numberFormatId, out string? formatCode);
            if (!isBuiltIn && numberFormatsById.TryGetValue(numberFormatId, out string? customCode)) {
                formatCode = customCode;
            }

            bool isDateLike = BiffBuiltInNumberFormat.IsDateLike(numberFormatId)
                || ExcelNumberFormatClassifier.LooksLikeDateFormat(formatCode);
            workbook.MutableCellFormats.Add(new LegacyXlsCellFormat(
                styleIndex,
                fontIndex,
                numberFormatId,
                isStyle,
                parentStyleIndex,
                applyNumberFormat,
                applyFont,
                fillPattern,
                fillForegroundColorIndex,
                fillBackgroundColorIndex,
                applyAlignment,
                horizontalAlignment,
                verticalAlignment,
                wrapText,
                textRotation,
                indent,
                shrinkToFit,
                readingOrder,
                applyProtection,
                locked,
                formulaHidden,
                quotePrefix,
                border,
                formatCode,
                isBuiltIn,
                isDateLike));
        }

        private static void ReadCellFormatApplyFlags(
            byte[] payload,
            out bool applyNumberFormat,
            out bool applyFont) {
            applyNumberFormat = false;
            applyFont = false;

            if (payload.Length < 10) {
                return;
            }

            ushort attributes = BiffRecordReader.ReadUInt16(payload, 8);
            applyNumberFormat = (attributes & 0x0400) != 0;
            applyFont = (attributes & 0x0800) != 0;
        }

        private static void ReadCellFormatProtection(
            byte[] payload,
            out bool applyProtection,
            out bool locked,
            out bool formulaHidden,
            out bool quotePrefix) {
            applyProtection = false;
            locked = true;
            formulaHidden = false;
            quotePrefix = false;

            if (payload.Length < 10) {
                return;
            }

            ushort protection = BiffRecordReader.ReadUInt16(payload, 4);
            ushort attributes = BiffRecordReader.ReadUInt16(payload, 8);
            locked = (protection & 0x0001) != 0;
            formulaHidden = (protection & 0x0002) != 0;
            quotePrefix = (protection & 0x0008) != 0;
            applyProtection = (attributes & 0x8000) != 0;
        }

        private static LegacyXlsBorder? ReadCellFormatBorder(byte[] payload) {
            if (payload.Length < 18) {
                return null;
            }

            ushort attributes = BiffRecordReader.ReadUInt16(payload, 8);
            if ((attributes & 0x2000) == 0) {
                return null;
            }

            uint sideBits = BiffRecordReader.ReadUInt32(payload, 10);
            uint topBottomBits = BiffRecordReader.ReadUInt32(payload, 14);
            byte leftStyle = (byte)(sideBits & 0x0f);
            byte rightStyle = (byte)((sideBits >> 4) & 0x0f);
            byte topStyle = (byte)((sideBits >> 8) & 0x0f);
            byte bottomStyle = (byte)((sideBits >> 12) & 0x0f);
            ushort leftColorIndex = (ushort)((sideBits >> 16) & 0x7f);
            ushort rightColorIndex = (ushort)((sideBits >> 23) & 0x7f);
            byte diagonalFlags = (byte)((sideBits >> 30) & 0x03);
            ushort topColorIndex = (ushort)(topBottomBits & 0x7f);
            ushort bottomColorIndex = (ushort)((topBottomBits >> 7) & 0x7f);
            ushort diagonalColorIndex = (ushort)((topBottomBits >> 14) & 0x7f);
            byte diagonalStyle = (byte)((topBottomBits >> 21) & 0x0f);

            bool hasBorder = leftStyle != 0
                || rightStyle != 0
                || topStyle != 0
                || bottomStyle != 0
                || (diagonalStyle != 0 && diagonalFlags != 0);
            if (!hasBorder) {
                return null;
            }

            return new LegacyXlsBorder(
                leftStyle,
                rightStyle,
                topStyle,
                bottomStyle,
                leftColorIndex,
                rightColorIndex,
                topColorIndex,
                bottomColorIndex,
                diagonalStyle,
                diagonalColorIndex,
                (diagonalFlags & 0x02) != 0,
                (diagonalFlags & 0x01) != 0);
        }

        private static void ReadCellFormatAlignment(
            byte[] payload,
            out bool applyAlignment,
            out byte horizontalAlignment,
            out byte verticalAlignment,
            out bool wrapText,
            out byte textRotation,
            out byte indent,
            out bool shrinkToFit,
            out byte readingOrder) {
            applyAlignment = false;
            horizontalAlignment = 0;
            verticalAlignment = 0;
            wrapText = false;
            textRotation = 0;
            indent = 0;
            shrinkToFit = false;
            readingOrder = 0;

            if (payload.Length < 10) {
                return;
            }

            byte alignment = payload[6];
            byte extendedAlignment = payload[8];
            ushort attributes = BiffRecordReader.ReadUInt16(payload, 8);
            applyAlignment = (attributes & 0x1000) != 0;
            horizontalAlignment = (byte)(alignment & 0x07);
            wrapText = (alignment & 0x08) != 0;
            verticalAlignment = (byte)((alignment >> 4) & 0x07);
            textRotation = payload[7];
            indent = (byte)(extendedAlignment & 0x0f);
            shrinkToFit = (extendedAlignment & 0x10) != 0;
            readingOrder = (byte)((extendedAlignment >> 6) & 0x03);
        }

        private static void ReadCellFormatFill(
            byte[] payload,
            out byte fillPattern,
            out ushort fillForegroundColorIndex,
            out ushort fillBackgroundColorIndex) {
            fillPattern = 0;
            fillForegroundColorIndex = 0;
            fillBackgroundColorIndex = 0;

            if (payload.Length < 20) {
                return;
            }

            uint fillPatternBits = BiffRecordReader.ReadUInt32(payload, 14);
            ushort fillColors = BiffRecordReader.ReadUInt16(payload, 18);
            fillPattern = (byte)((fillPatternBits >> 26) & 0x3f);
            fillForegroundColorIndex = (ushort)(fillColors & 0x7f);
            fillBackgroundColorIndex = (ushort)((fillColors >> 7) & 0x7f);
        }

        private static LegacyXlsWorksheet? TryReadBoundSheet(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics) {
            try {
                if (record.Payload.Length < 8) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-BOUNDSHEET-SHORT",
                        "A BoundSheet8 record was too short to parse.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return null;
                }

                int streamOffset = checked((int)BiffRecordReader.ReadUInt32(record.Payload, 0));
                byte visibility = record.Payload[4];
                byte sheetType = record.Payload[5];
                int nameOffset = 6;
                string name = BiffStringReader.ReadShortUnicodeString(record.Payload, ref nameOffset);
                return new LegacyXlsWorksheet(name, streamOffset, visibility, sheetType);
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-BOUNDSHEET-INVALID",
                    $"A BoundSheet8 record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return null;
            }
        }

        private static void MoveDialogSheetsToUnsupported(byte[] workbookStream, LegacyXlsWorkbook workbook, LegacyXlsImportOptions options) {
            for (int i = workbook.MutableWorksheets.Count - 1; i >= 0; i--) {
                LegacyXlsWorksheet sheet = workbook.MutableWorksheets[i];
                if (!TryReadDialogSheetFlag(workbookStream, sheet, out bool isDialogSheet, out int wsBoolOffset) || !isDialogSheet) {
                    continue;
                }

                workbook.MutableWorksheets.RemoveAt(i);
                LegacyXlsUnsupportedSheet unsupportedSheet = ToUnsupportedSheet(sheet, LegacyXlsUnsupportedSheetKind.DialogSheet);
                workbook.MutableUnsupportedSheets.Add(unsupportedSheet);
                workbook.MutableUnsupportedFeatures.Add(BiffUnsupportedRecordDiagnostics.CreateUnsupportedDialogSheetFeature(wsBoolOffset, unsupportedSheet));
                if (options.ReportUnsupportedRecords) {
                    BiffUnsupportedRecordDiagnostics.AddUnsupportedDialogSheetDiagnostic(workbook.MutableDiagnostics, wsBoolOffset, unsupportedSheet);
                }
            }
        }

        private static bool TryReadDialogSheetFlag(byte[] workbookStream, LegacyXlsWorksheet sheet, out bool isDialogSheet, out int wsBoolOffset) {
            isDialogSheet = false;
            wsBoolOffset = -1;
            if (sheet.StreamOffset < 0 || sheet.StreamOffset >= workbookStream.Length) {
                return false;
            }

            int offset = sheet.StreamOffset;
            while (offset + 4 <= workbookStream.Length) {
                ushort type = BiffRecordReader.ReadUInt16(workbookStream, offset);
                ushort length = BiffRecordReader.ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) {
                    return false;
                }

                if (type == (ushort)BiffRecordType.Eof) {
                    return false;
                }

                if (type == (ushort)BiffRecordType.WsBool) {
                    wsBoolOffset = offset;
                    isDialogSheet = length >= 2 && (BiffRecordReader.ReadUInt16(workbookStream, payloadOffset) & 0x0010) != 0;
                    return true;
                }

                offset = payloadOffset + length;
            }

            return false;
        }

        private static LegacyXlsUnsupportedSheet ToUnsupportedSheet(LegacyXlsWorksheet sheet, LegacyXlsUnsupportedSheetKind kind) {
            return new LegacyXlsUnsupportedSheet(sheet.Name, sheet.StreamOffset, sheet.Visibility, sheet.SheetType, kind);
        }

        private static LegacyXlsUnsupportedSheetKind ToUnsupportedSheetKind(byte sheetType) {
            switch (sheetType) {
                case 0x01:
                    return LegacyXlsUnsupportedSheetKind.MacroSheet;
                case 0x02:
                    return LegacyXlsUnsupportedSheetKind.ChartSheet;
                case 0x06:
                    return LegacyXlsUnsupportedSheetKind.VbaModuleSheet;
                default:
                    return LegacyXlsUnsupportedSheetKind.Unknown;
            }
        }
    }
}
