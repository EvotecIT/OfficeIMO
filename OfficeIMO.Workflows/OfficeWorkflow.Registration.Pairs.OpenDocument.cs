using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        private async Task<OfficeConversionReport> ConvertOpenDocumentAsync(Stream input, Stream output, OfficeConversionOptions options,
            OfficeWorkflowRouteSettings settings, CancellationToken token) {
            OdfLoadOptions load = settings.OpenDocumentLoad ?? new();
            load.Password ??= options.SourcePassword;
            load.MaxPackageBytes = Math.Min(load.MaxPackageBytes, options.MaximumInputBytes);
            load.MaxXmlCharacters = Math.Min(load.MaxXmlCharacters, options.MaximumXmlCharactersInPart);
            WordOpenDocumentConversionOptions word = settings.WordOpenDocument ?? new();
            ExcelOpenDocumentConversionOptions excel = settings.ExcelOpenDocument ?? new();
            PowerPointOpenDocumentConversionOptions powerPoint = settings.PowerPointOpenDocument ?? new();
            if (options.FailOnLoss) {
                word.LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss;
                excel.LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss;
                powerPoint.LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss;
            }
            string? sourceExtension = Path.GetExtension(options.SourceName)?.ToLowerInvariant();
            switch (Id) {
                case "docx-odt": {
                    WordLoadOptions reading = OfficeConverter.WordInputOptions(options);
                    using WordDocument document = await WordDocument.LoadAsync(input, reading, token).ConfigureAwait(false);
                    return await RetainSourceAsync(document.CreateLoadReport(options.SourceName!, TimeSpan.Zero),
                        async () => await SaveOpenDocumentAsync(await document.ToOpenDocumentAsync(word, token).ConfigureAwait(false), output,
                            settings.OpenDocumentSave, options.FailOnLoss, token).ConfigureAwait(false), options.FailOnLoss).ConfigureAwait(false);
                }
                case "xlsx-ods": {
                    ExcelLoadOptions reading = OfficeConverter.ExcelInputOptions(options);
                    using ExcelDocument document = await ExcelDocument.LoadAsync(input, reading, token).ConfigureAwait(false);
                    return await RetainSourceAsync(document.CreateLoadReport(options.SourceName!, TimeSpan.Zero),
                        async () => await SaveOpenDocumentAsync(await document.ToOpenDocumentAsync(excel, token).ConfigureAwait(false), output,
                            settings.OpenDocumentSave, options.FailOnLoss, token).ConfigureAwait(false), options.FailOnLoss).ConfigureAwait(false);
                }
                case "pptx-odp": {
                    PowerPointLoadOptions reading = OfficeConverter.PowerPointInputOptions(options);
                    using PowerPointDocument document = await PowerPointDocument.LoadAsync(input, reading, token).ConfigureAwait(false);
                    return await RetainSourceAsync(document.CreateLoadReport(options.SourceName!, TimeSpan.Zero),
                        async () => await SaveOpenDocumentAsync(await document.ToOpenDocumentAsync(powerPoint, token).ConfigureAwait(false), output,
                            settings.OpenDocumentSave, options.FailOnLoss, token).ConfigureAwait(false), options.FailOnLoss).ConfigureAwait(false);
                }
                case "odt-docx": {
                    using OdtDocument document = sourceExtension == ".fodt"
                        ? OdtDocument.LoadFlatXml(input, load)
                        : await OdtDocument.LoadAsync(input, load, token).ConfigureAwait(false);
                    return await SaveModelAsync(await document.ToWordAsync(word, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, null, token), token).ConfigureAwait(false);
                }
                case "ods-xlsx": {
                    using OdsDocument document = sourceExtension == ".fods"
                        ? OdsDocument.LoadFlatXml(input, load)
                        : await OdsDocument.LoadAsync(input, load, token).ConfigureAwait(false);
                    return await SaveModelAsync(await document.ToExcelAsync(excel, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, null, token), token).ConfigureAwait(false);
                }
                case "odp-pptx": {
                    using OdpPresentation document = sourceExtension == ".fodp"
                        ? OdpPresentation.LoadFlatXml(input, load)
                        : await OdpPresentation.LoadAsync(input, load, token).ConfigureAwait(false);
                    return await SaveModelAsync(await document.ToPowerPointAsync(powerPoint, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, null, token), token).ConfigureAwait(false);
                }
                case "odt-pdf": {
                    using OdtDocument document = sourceExtension == ".fodt"
                        ? OdtDocument.LoadFlatXml(input, load)
                        : await OdtDocument.LoadAsync(input, load, token).ConfigureAwait(false);
                    OdtPdfExportOptions pdf = settings.OdtPdf ?? new(); pdf.FailOnLoss |= options.FailOnLoss;
                    return await document.SaveAsPdfAsync(output, pdf, token).ConfigureAwait(false);
                }
                case "ods-pdf": {
                    using OdsDocument document = sourceExtension == ".fods"
                        ? OdsDocument.LoadFlatXml(input, load)
                        : await OdsDocument.LoadAsync(input, load, token).ConfigureAwait(false);
                    OdsPdfExportOptions pdf = settings.OdsPdf ?? new(); pdf.FailOnLoss |= options.FailOnLoss;
                    return await document.SaveAsPdfAsync(output, pdf, token).ConfigureAwait(false);
                }
                default: {
                    using OdpPresentation document = sourceExtension == ".fodp"
                        ? OdpPresentation.LoadFlatXml(input, load)
                        : await OdpPresentation.LoadAsync(input, load, token).ConfigureAwait(false);
                    OdpPdfExportOptions pdf = settings.OdpPdf ?? new(); pdf.FailOnLoss |= options.FailOnLoss;
                    return await document.SaveAsPdfAsync(output, pdf, token).ConfigureAwait(false);
                }
            }
        }

        private static async Task<OfficeConversionReport> SaveOpenDocumentAsync<T>(OfficeConversionResult<T> converted,
            Stream output, OdfSaveOptions? options, bool failOnLoss, CancellationToken token) where T : OdfDocument {
            using T document = converted.Value;
            OdfSaveResult? saved = null;
            await OfficeConversionFailures.RunAsync(converted.Report, async () => {
                saved = await document.SaveAsync(output, options, token).ConfigureAwait(false);
            }).ConfigureAwait(false);
            OfficeConversionReport report = Combine(converted.Report, saved!.Report);
            if (failOnLoss) report.ThrowIfLoss();
            return report;
        }
    }
}
