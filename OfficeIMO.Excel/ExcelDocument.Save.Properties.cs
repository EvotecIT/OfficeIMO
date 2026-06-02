using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Shared;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        private sealed class PackagePropertiesSnapshot {
            private readonly string? _title;
            private readonly string? _creator;
            private readonly string? _subject;
            private readonly string? _category;
            private readonly string? _description;
            private readonly string? _keywords;
            private readonly string? _lastModifiedBy;
            private readonly string? _version;
            private readonly DateTime? _created;
            private readonly DateTime? _modified;
            private readonly DateTime? _lastPrinted;

            private PackagePropertiesSnapshot(
                string? title,
                string? creator,
                string? subject,
                string? category,
                string? description,
                string? keywords,
                string? lastModifiedBy,
                string? version,
                DateTime? created,
                DateTime? modified,
                DateTime? lastPrinted) {
                _title = title;
                _creator = creator;
                _subject = subject;
                _category = category;
                _description = description;
                _keywords = keywords;
                _lastModifiedBy = lastModifiedBy;
                _version = version;
                _created = created;
                _modified = modified;
                _lastPrinted = lastPrinted;
            }

            public static PackagePropertiesSnapshot Capture(SpreadsheetDocument document) {
                try {
                    var props = document.PackageProperties;
                    return new PackagePropertiesSnapshot(
                        props.Title,
                        props.Creator,
                        props.Subject,
                        props.Category,
                        props.Description,
                        props.Keywords,
                        props.LastModifiedBy,
                        props.Version,
                        props.Created,
                        props.Modified,
                        props.LastPrinted);
                } catch {
                    return new PackagePropertiesSnapshot(null, null, null, null, null, null, null, null, null, null, null);
                }
            }

            public void ApplyTo(string packagePath) {
                if (string.IsNullOrWhiteSpace(packagePath) || !File.Exists(packagePath)) {
                    return;
                }

                try {
                    using var package = Package.Open(packagePath, FileMode.Open, FileAccess.ReadWrite);
                    var dst = package.PackageProperties;
                    dst.Title = _title;
                    dst.Creator = _creator;
                    dst.Subject = _subject;
                    dst.Category = _category;
                    dst.Description = _description;
                    dst.Keywords = _keywords;
                    dst.LastModifiedBy = _lastModifiedBy;
                    dst.Version = _version;
                    dst.Created = _created;
                    dst.Modified = _modified ?? DateTime.UtcNow;
                    dst.LastPrinted = _lastPrinted;
                } catch {
                }
            }

            public byte[] ApplyTo(byte[] packageBytes) {
                if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
                if (packageBytes.Length == 0) return packageBytes;

                try {
                    using var working = new MemoryStream(packageBytes.Length + StreamBufferSize);
                    working.Write(packageBytes, 0, packageBytes.Length);
                    working.Position = 0;

                    using (var package = Package.Open(working, FileMode.Open, FileAccess.ReadWrite)) {
                        var dst = package.PackageProperties;
                        dst.Title = _title;
                        dst.Creator = _creator;
                        dst.Subject = _subject;
                        dst.Category = _category;
                        dst.Description = _description;
                        dst.Keywords = _keywords;
                        dst.LastModifiedBy = _lastModifiedBy;
                        dst.Version = _version;
                        dst.Created = _created;
                        dst.Modified = _modified ?? DateTime.UtcNow;
                        dst.LastPrinted = _lastPrinted;
                    }

                    if (working.CanSeek) {
                        working.Position = 0;
                    }

                    return working.ToArray();
                } catch {
                    return packageBytes;
                }
            }
        }

        private sealed class SavePayload {
            public SavePayload(byte[] packageBytes, PackagePropertiesSnapshot properties, bool documentClosed, bool normalizeContentTypes, bool applyPackageProperties) {
                PackageBytes = packageBytes;
                Properties = properties;
                DocumentClosed = documentClosed;
                NormalizeContentTypes = normalizeContentTypes;
                ApplyPackageProperties = applyPackageProperties;
            }

            public byte[] PackageBytes { get; }
            public PackagePropertiesSnapshot Properties { get; }
            public bool DocumentClosed { get; }
            public bool NormalizeContentTypes { get; }
            public bool ApplyPackageProperties { get; }
        }
    }
}
