using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static void ValidateDestinationSaveFormats(OfficeConversionOptions settings) {
        OfficeDocumentFormat target = FormatForExtension(Path.GetExtension(settings.TargetName));
        if (target == OfficeDocumentFormat.Word && settings.WordSave is { Format: not WordFileFormat.Docx }
            || target == OfficeDocumentFormat.Excel && settings.ExcelSave is { Format: not ExcelFileFormat.Xlsx }
            || target == OfficeDocumentFormat.PowerPoint && settings.PowerPointSave is { Format: not PowerPointFileFormat.Pptx })
            throw new ArgumentException("Cross-format save options must retain the physical format selected by the destination filename.", nameof(settings));
    }

    private static async Task ApplyRegisteredSaveOptionsAsync(Stream output, OfficeConversionOptions settings,
        OfficeConversionReport report, CancellationToken token) {
        OfficeDocumentFormat target = FormatForExtension(Path.GetExtension(settings.TargetName));
        if (!(target == OfficeDocumentFormat.Word && settings.WordSave != null
            || target == OfficeDocumentFormat.Excel && settings.ExcelSave != null
            || target == OfficeDocumentFormat.PowerPoint && settings.PowerPointSave != null)) return;
        await OfficeConversionFailures.RunAsync(report, async () => {
            output.Position = 0;
            bool zip = output.ReadByte() == 'P' && output.ReadByte() == 'K';
            output.Position = 0;
            using var saved = new OfficeBoundedMemoryStream(settings.MaximumOutputBytes);
            switch (target) {
                case OfficeDocumentFormat.Word:
                    using (var model = await WordDocument.LoadAsync(output, new WordLoadOptions {
                        MaxInputBytes = settings.MaximumOutputBytes, Password = zip ? null : settings.WordSave!.Password
                    }, token).ConfigureAwait(false))
                        await model.SaveAsync(saved, settings.WordSave, token).ConfigureAwait(false);
                    break;
                case OfficeDocumentFormat.Excel:
                    using (var model = await ExcelDocument.LoadAsync(output, new ExcelLoadOptions {
                        MaxInputBytes = settings.MaximumOutputBytes, Password = zip ? null : settings.ExcelSave!.Password
                    }, token).ConfigureAwait(false))
                        await model.SaveAsync(saved, settings.ExcelSave, token).ConfigureAwait(false);
                    break;
                case OfficeDocumentFormat.PowerPoint:
                    using (var model = await PowerPointPresentation.LoadAsync(output, new PowerPointLoadOptions {
                        MaxInputBytes = settings.MaximumOutputBytes, Password = zip ? null : settings.PowerPointSave!.Password
                    }, token).ConfigureAwait(false))
                        await model.SaveAsync(saved, settings.PowerPointSave, token).ConfigureAwait(false);
                    break;
            }
            token.ThrowIfCancellationRequested();
            output.SetLength(0); output.Position = 0; saved.Position = 0;
            await saved.CopyToAsync(output, 81920, token).ConfigureAwait(false);
        }).ConfigureAwait(false);
    }
}
