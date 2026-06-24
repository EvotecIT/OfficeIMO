using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for Excel write-reservation metadata. This is not package encryption and does not protect worksheet or workbook structure.
    /// </summary>
    public sealed class ExcelWorkbookWriteReservationOptions {
        /// <summary>
        /// Suggests opening the workbook as read-only.
        /// </summary>
        public bool ReadOnlyRecommended { get; set; }

        /// <summary>
        /// Optional user name displayed by Excel-compatible applications for the write reservation.
        /// </summary>
        public string? UserName { get; set; }

        /// <summary>
        /// Optional write-reservation password. This uses Excel's legacy reservation hash and is not encryption.
        /// </summary>
        public string? Password { get; set; }

        /// <summary>
        /// Optional precomputed legacy write-reservation hash. When set, this value is written as-is.
        /// </summary>
        public string? LegacyPasswordHash { get; set; }
    }

    /// <summary>
    /// Snapshot of workbook write-reservation metadata.
    /// </summary>
    public sealed class ExcelWorkbookWriteReservationInfo {
        internal ExcelWorkbookWriteReservationInfo(bool exists, bool readOnlyRecommended, string? userName, string? legacyPasswordHash) {
            Exists = exists;
            ReadOnlyRecommended = readOnlyRecommended;
            UserName = userName;
            LegacyPasswordHash = legacyPasswordHash;
        }

        /// <summary>Gets a value indicating whether the workbook contains a file-sharing/write-reservation node.</summary>
        public bool Exists { get; }

        /// <summary>Gets a value indicating whether Excel-compatible applications should recommend read-only opening.</summary>
        public bool ReadOnlyRecommended { get; }

        /// <summary>Gets the reservation user name when present.</summary>
        public string? UserName { get; }

        /// <summary>Gets the legacy reservation password hash when present.</summary>
        public string? LegacyPasswordHash { get; }

        /// <summary>Gets a value indicating whether a write-reservation password hash is present.</summary>
        public bool HasPasswordHash => !string.IsNullOrWhiteSpace(LegacyPasswordHash);
    }

    public partial class ExcelDocument {
        /// <summary>
        /// Gets workbook write-reservation metadata. This is separate from workbook protection and package encryption.
        /// </summary>
        public ExcelWorkbookWriteReservationInfo GetWriteReservation() {
            FileSharing? fileSharing = WorkbookRoot.GetFirstChild<FileSharing>();
            if (fileSharing == null) {
                return new ExcelWorkbookWriteReservationInfo(false, false, null, null);
            }

            return new ExcelWorkbookWriteReservationInfo(
                exists: true,
                readOnlyRecommended: fileSharing.ReadOnlyRecommended?.Value ?? false,
                userName: fileSharing.UserName?.Value,
                legacyPasswordHash: GetFileSharingReservationPassword(fileSharing));
        }

        /// <summary>
        /// Writes workbook write-reservation metadata. This is not package encryption and does not protect worksheet or workbook structure.
        /// </summary>
        public void SetWriteReservation(ExcelWorkbookWriteReservationOptions? options = null) {
            var opts = options ?? new ExcelWorkbookWriteReservationOptions { ReadOnlyRecommended = true };
            FileSharing fileSharing = WorkbookRoot.GetFirstChild<FileSharing>() ?? new FileSharing();
            if (fileSharing.Parent == null) {
                InsertFileSharingInSchemaOrder(WorkbookRoot, fileSharing);
            }

            fileSharing.ReadOnlyRecommended = opts.ReadOnlyRecommended ? true : (bool?)null;
            if (string.IsNullOrWhiteSpace(opts.UserName)) {
                fileSharing.UserName = null;
                fileSharing.RemoveAttribute("userName", string.Empty);
            } else {
                fileSharing.UserName = opts.UserName;
            }

            string? hash = ExcelProtectionHash.ResolveLegacyHash(opts.Password, opts.LegacyPasswordHash);
            if (hash == null) {
                fileSharing.RemoveAttribute("reservationPassword", string.Empty);
            } else {
                fileSharing.SetAttribute(new OpenXmlAttribute("reservationPassword", string.Empty, hash));
            }

            WorkbookRoot.Save();
            MarkPackageDirty();
        }

        /// <summary>
        /// Removes workbook write-reservation metadata.
        /// </summary>
        public void ClearWriteReservation() {
            FileSharing? fileSharing = WorkbookRoot.GetFirstChild<FileSharing>();
            if (fileSharing == null) {
                return;
            }

            WorkbookRoot.RemoveChild(fileSharing);
            WorkbookRoot.Save();
            MarkPackageDirty();
        }

        private static string? GetFileSharingReservationPassword(FileSharing fileSharing) {
            OpenXmlAttribute attribute = fileSharing.GetAttribute("reservationPassword", string.Empty);
            return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
        }

        private static void InsertFileSharingInSchemaOrder(Workbook workbook, FileSharing fileSharing) {
            OpenXmlElement? before = workbook.GetFirstChild<WorkbookProperties>();
            before ??= workbook.GetFirstChild<WorkbookProtection>();
            before ??= workbook.GetFirstChild<BookViews>();
            before ??= workbook.GetFirstChild<Sheets>();

            if (before != null) {
                workbook.InsertBefore(fileSharing, before);
            } else if (workbook.GetFirstChild<FileVersion>() is FileVersion fileVersion) {
                workbook.InsertAfter(fileSharing, fileVersion);
            } else {
                workbook.InsertAt(fileSharing, 0);
            }
        }
    }
}
