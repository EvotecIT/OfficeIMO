using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Core.Internal;
using System.IO.Packaging;
using System.Xml;
using System.Diagnostics;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        private static async Task<byte[]> ReadAllBytesCompatAsync(string path, CancellationToken ct,
            ExcelLoadOptions options) {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete, 8192, FileOptions.Asynchronous)) {
                return await ReadAllBytesAsync(fs, ct, options).ConfigureAwait(false);
            }
        }

        private static OpenSettings CreateOpenSettings(OfficeOpenXmlLoadSettings? openSettings) =>
            openSettings.ToOpenXml();
    }
}
