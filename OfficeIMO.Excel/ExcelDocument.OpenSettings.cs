using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Core.Internal;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        private static async Task<byte[]> ReadAllBytesCompatAsync(string path, CancellationToken ct,
            ExcelLoadOptions options) {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete, 8192, FileOptions.Asynchronous)) {
                long? inputLimit = ResolveInputLimit(options);
                return options.PackageSecurity != null
                    && inputLimit == options.PackageSecurity.MaxPackageBytes
                    ? await OfficePackageSecurityInspector.ReadBoundedAsync(fs, options.PackageSecurity, ct)
                        .ConfigureAwait(false)
                    : await OfficeStreamReader.ReadAllBytesAsync(fs, ct, inputLimit).ConfigureAwait(false);
            }
        }

        private static OpenSettings CreateOpenSettings(OfficeOpenXmlLoadSettings? openSettings) =>
            openSettings.ToOpenXml();
    }
}
