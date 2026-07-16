using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing.Internal;
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
            OfficePackageSecurityOptions? securityOptions = null) {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete, 8192, FileOptions.Asynchronous)) {
                return securityOptions == null
                    ? await OfficeStreamReader.ReadAllBytesAsync(fs, ct).ConfigureAwait(false)
                    : await OfficePackageSecurityInspector.ReadBoundedAsync(fs, securityOptions, ct)
                        .ConfigureAwait(false);
            }
        }

        private static OpenSettings CreateOpenSettings(OpenSettings? openSettings) {
            if (openSettings is null) {
                return new OpenSettings { AutoSave = false };
            }

            return new OpenSettings {
                AutoSave = false,
                CompatibilityLevel = openSettings.CompatibilityLevel,
                MarkupCompatibilityProcessSettings = openSettings.MarkupCompatibilityProcessSettings,
                MaxCharactersInPart = openSettings.MaxCharactersInPart,
            };
        }
    }
}
