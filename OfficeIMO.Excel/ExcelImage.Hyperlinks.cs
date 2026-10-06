using DocumentFormat.OpenXml.Packaging;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelImage {
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// Gets or sets the picture's external hyperlink target. Null removes the click link.
        /// Targets are stored in the workbook and are never opened by OfficeIMO.
        /// </summary>
        /// <remarks>
        /// A save that reloads the package invalidates image handles. Obtain current images from
        /// <see cref="ExcelSheet.Images"/> after such saves or removing a worksheet.
        /// Reading is supported on read-only workbooks; changing the link requires a writable workbook.
        /// </remarks>
        public Uri? HyperlinkUri {
            get => Locking.ExecuteRead(_document.EnsureLock(), () => {
                RequireCurrentHyperlinkOwner();
                string? id = DrawingProperties?.GetFirstChild<A.HyperlinkOnClick>()?.Id?.Value;
                return string.IsNullOrEmpty(id)
                    ? null
                    : _drawingsPart.HyperlinkRelationships.FirstOrDefault(item => item.Id == id && item.IsExternal)?.Uri;
            });
            set => Locking.ExecuteWrite(_document.EnsureLock(), () => {
                RequireCurrentHyperlinkOwner();
                if (_document._spreadSheetDocument.FileOpenAccess == FileAccess.Read) {
                    throw new InvalidOperationException("Picture hyperlinks cannot be changed in a read-only workbook.");
                }
                Xdr.NonVisualDrawingProperties properties = DrawingProperties
                    ?? throw new InvalidOperationException("The picture is missing its drawing properties.");
                A.HyperlinkOnClick? link = properties.GetFirstChild<A.HyperlinkOnClick>();
                string? previousId = link?.Id?.Value;
                if (value == null) {
                    link?.Remove();
                } else {
                    HyperlinkRelationship relationship = _drawingsPart.HyperlinkRelationships
                        .FirstOrDefault(item => item.IsExternal && string.Equals(item.Uri.OriginalString, value.OriginalString, StringComparison.Ordinal))
                        ?? _drawingsPart.AddHyperlinkRelationship(value, true);
                    if (link == null) {
                        link = new A.HyperlinkOnClick();
                        properties.AddChild(link, true);
                    }
                    link.Id = relationship.Id;
                    link.Action = null;
                }
                RemoveUnusedPictureHyperlink(previousId);
                Save();
            });
        }

        private void RequireCurrentHyperlinkOwner() {
            WorkbookPart? workbook = _document._spreadSheetDocument.WorkbookPart;
            if (workbook == null
                || !workbook.WorksheetParts.Any(part => ReferenceEquals(part.DrawingsPart, _drawingsPart))
                || !ReferenceEquals(_picture.Ancestors<Xdr.WorksheetDrawing>().FirstOrDefault(), _drawingsPart.WorksheetDrawing)) {
                throw new InvalidOperationException(
                    "The picture is no longer part of the current workbook package. Obtain a current image from ExcelDocument.Sheets before reading or changing its hyperlink.");
            }
        }

        private void RemoveUnusedPictureHyperlink(string? id) {
            if (string.IsNullOrEmpty(id)) return;
            HyperlinkRelationship? relationship = _drawingsPart.HyperlinkRelationships.FirstOrDefault(item => item.Id == id);
            if (relationship == null) return;
            bool referenced = _drawingsPart.WorksheetDrawing!.Descendants().Any(element =>
                element.GetAttributes().Any(attribute => attribute.NamespaceUri == RelationshipNamespace
                    && attribute.LocalName == "id" && attribute.Value == id));
            if (!referenced) _drawingsPart.DeleteReferenceRelationship(relationship);
        }
    }
}
