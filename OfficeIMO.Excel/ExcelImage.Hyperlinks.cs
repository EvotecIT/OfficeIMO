using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelImage {
        /// <summary>
        /// Gets or sets the picture's external hyperlink target. Null removes the click link.
        /// Targets are stored in the workbook and are never opened by OfficeIMO.
        /// </summary>
        /// <remarks>
        /// Image handles remain usable after saving while the picture belongs to the workbook.
        /// Removing the picture or its worksheet invalidates its handles.
        /// Reading is supported on read-only workbooks; changing the link requires a writable workbook.
        /// </remarks>
        public Uri? HyperlinkUri {
            get => Locking.ExecuteRead(_document.EnsureLock(), () => {
                RequireCurrentHyperlinkOwner();
                string? id = DrawingProperties.Child(A, "hlinkClick").RelationshipAttribute("id");
                return string.IsNullOrEmpty(id)
                    ? null
                    : DrawingsPartRef.GetHyperlinkUri(id!, externalOnly: true);
            });
            set => Locking.ExecuteWrite(_document.EnsureLock(), () => {
                RequireCurrentHyperlinkOwner();
                if (ExcelPartRef.IsReadOnlyPackage(_document)) {
                    throw new InvalidOperationException("Picture hyperlinks cannot be changed in a read-only workbook.");
                }
                OfficeXmlNode properties = DrawingProperties;
                if (properties.IsNone) throw new InvalidOperationException("The picture is missing its drawing properties.");
                OfficeXmlNode link = properties.Child(A, "hlinkClick");
                string? previousId = link.RelationshipAttribute("id");
                if (value == null) {
                    if (link.Exists) link.Remove();
                } else {
                    string relationshipId = DrawingsPartRef.GetOrAddExternalHyperlinkWithSameText(value);
                    if (link.IsNone) {
                        link = properties.AddOrdered(properties.New(A, "hlinkClick"), ExcelDrawingXml.NonVisualDrawingPropertiesOrder);
                    }
                    link.SetAttribute(R, "id", relationshipId);
                    link.SetAttribute("action", null);
                }
                RemoveUnusedPictureHyperlink(previousId);
                Save();
            });
        }

        private void RequireCurrentHyperlinkOwner() {
            ExcelPartRef drawingsPart = DrawingsPartRef;
            if (!ExcelPartRef.WorksheetParts(_document).Any(part => part.Drawings == drawingsPart)
                || PictureNode.Ancestor(Xdr, "wsDr") != drawingsPart.Root) {
                throw new InvalidOperationException(
                    "The picture is no longer part of the current workbook package. Obtain a current image from ExcelDocument.Sheets before reading or changing its hyperlink.");
            }
        }

        private void RemoveUnusedPictureHyperlink(string? id) {
            if (string.IsNullOrEmpty(id)) return;
            ExcelPartRef drawingsPart = DrawingsPartRef;
            if (!drawingsPart.HasHyperlink(id!)) return;
            bool referenced = ExcelDrawingXml.HasRelationshipAttribute(drawingsPart.Root, "id", id!);
            if (!referenced) drawingsPart.RemoveHyperlink(id!);
        }
    }
}
