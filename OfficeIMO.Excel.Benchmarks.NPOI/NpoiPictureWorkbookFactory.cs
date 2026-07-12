using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;

internal static class NpoiPictureWorkbookFactory {
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");

    internal static int GetPictureCount(int rowCount) {
        return Math.Min(Math.Max(rowCount / 250, 1), 12);
    }

    internal static byte[] WriteHssfPictureWorkbook(int pictureCount) {
        if (pictureCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pictureCount));
        }

        using var stream = new MemoryStream();
        using var workbook = new HSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Pictures");

        IRow header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("PictureId");
        header.CreateCell(1).SetCellValue("Kind");

        HSSFPatriarch drawing = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
        for (int i = 0; i < pictureCount; i++) {
            int rowIndex = i + 1;
            IRow row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue(i + 1);
            row.CreateCell(1).SetCellValue("Png");

            int pictureIndex = workbook.AddPicture(PngBytes, PictureType.PNG);
            var anchor = new HSSFClientAnchor(0, 0, 512, 128, 2, rowIndex, 4, rowIndex + 2);
            drawing.CreatePicture(anchor, pictureIndex);
        }

        workbook.Write(stream, leaveOpen: true);
        return stream.ToArray();
    }
}

internal static class NpoiPictureComparison {
    internal static int ReadOfficeImoXlsPictures(byte[] workbookBytes, int expectedPictureCount, Func<int, object?, int> addValueMetric) {
        LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedContent = true });
        List<LegacyXlsDrawingBlipStoreEntry> blipEntries = workbook.DrawingRecords
            .SelectMany(record => record.BlipStoreEntries)
            .ToList();
        int pictureObjectCount = workbook.DrawingRecords.Count(record => record.ObjectTypeKind == LegacyXlsDrawingObjectType.Picture);
        int pictureFrameCount = workbook.DrawingRecords
            .SelectMany(record => record.ShapeEntries)
            .Count(shape => string.Equals(shape.ShapeTypeName, "PictureFrame", StringComparison.Ordinal));

        ValidatePictureCounts(blipEntries.Count, pictureObjectCount, expectedPictureCount, "OfficeIMO legacy XLS");

        int metric = 0;
        metric = addValueMetric(metric, blipEntries.Count);
        metric = addValueMetric(metric, pictureFrameCount);
        foreach (LegacyXlsDrawingBlipStoreEntry entry in blipEntries.OrderBy(entry => entry.EmbeddedBlipPayloadSha256, StringComparer.Ordinal)) {
            metric = addValueMetric(metric, entry.RecordInstanceBlipTypeName);
            metric = addValueMetric(metric, entry.EmbeddedBlipPayloadAvailableLength ?? 0);
            metric = addValueMetric(metric, entry.ReferenceCount ?? 0);
        }

        return metric;
    }

    internal static int ReadNpoiWorkbookPictures(byte[] workbookBytes, int expectedPictureCount, Func<int, object?, int> addValueMetric) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new HSSFWorkbook(stream);
        List<IPictureData> pictures = workbook.GetAllPictures().Cast<IPictureData>().ToList();
        ValidatePictureCounts(pictures.Count, pictures.Count, expectedPictureCount, "NPOI HSSF");

        int metric = 0;
        metric = addValueMetric(metric, pictures.Count);
        foreach (IPictureData picture in pictures.OrderBy(picture => picture.Data.Length)) {
            metric = addValueMetric(metric, picture.PictureType.ToString());
            metric = addValueMetric(metric, picture.Data.Length);
        }

        return metric;
    }

    private static void ValidatePictureCounts(int blipOrPictureCount, int pictureObjectCount, int expectedPictureCount, string libraryName) {
        if (blipOrPictureCount != expectedPictureCount
            || pictureObjectCount != expectedPictureCount) {
            throw new InvalidOperationException(
                $"{libraryName} picture counts did not match. "
                + $"Images {blipOrPictureCount}/{expectedPictureCount}, "
                + $"Objects {pictureObjectCount}/{expectedPictureCount}.");
        }
    }
}
