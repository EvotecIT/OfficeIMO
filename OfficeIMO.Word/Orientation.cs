using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Contains methods to change print orientation for existing documents.
    /// </summary>
    internal class OrientationOfPage {
        // https://github.com/OfficeDev/open-xml-docs/blob/master/docs/how-to-change-the-print-orientation-of-a-word-processing-document.md
        /// <summary>
        /// Changes the print orientation of an existing Word document.
        /// </summary>
        /// <param name="fileName">Path to the document.</param>
        /// <param name="newOrientation">Desired page orientation.</param>
        public static void SetPrintOrientation(string fileName, PageOrientationValues newOrientation) {
            using (var document = WordprocessingDocument.Open(fileName, true)) {
                bool documentChanged = false;

                var docPart = document.MainDocumentPart;
                var sections = docPart.Document.Descendants<SectionProperties>();

                foreach (SectionProperties sectPr in sections) {
                    bool pageOrientationChanged = false;

                    PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                    if (pgSz != null) {
                        // No Orient property? Create it now. Otherwise, just
                        // set its value. Assume that the default orientation
                        // is Portrait.
                        if (pgSz.Orient == null) {
                            // Need to create the attribute. You do not need to
                            // create the Orient property if the property does not
                            // already exist, and you are setting it to Portrait.
                            // That is the default value.
                            if (newOrientation != PageOrientationValues.Portrait) {
                                pageOrientationChanged = true;
                                documentChanged = true;
                                pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
                            }
                        } else {
                            // The Orient property exists, but its value
                            // is different than the new value.
                            if (pgSz.Orient.Value != newOrientation) {
                                pgSz.Orient.Value = newOrientation;
                                pageOrientationChanged = true;
                                documentChanged = true;
                            }
                        }

                        if (pageOrientationChanged) {
                            // Changing the orientation is not enough. You must also
                            // change the page size.
                            var width = pgSz.Width;
                            var height = pgSz.Height;
                            pgSz.Width = height;
                            pgSz.Height = width;

                            PageMargin pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();
                            if (pgMar != null) {
                                // Rotate margins. Printer settings control how far you
                                // rotate when switching to landscape mode. Not having those
                                // settings, this code rotates 90 degrees. You could easily
                                // modify this behavior, or make it a parameter for the
                                // procedure.
                                var top = pgMar.Top.Value;
                                var bottom = pgMar.Bottom.Value;
                                var left = pgMar.Left.Value;
                                var right = pgMar.Right.Value;

                                pgMar.Top = new Int32Value((int)left);
                                pgMar.Bottom = new Int32Value((int)right);
                                pgMar.Left = new UInt32Value((uint)System.Math.Max(0, bottom));
                                pgMar.Right = new UInt32Value((uint)System.Math.Max(0, top));
                            }
                        }
                    }
                }

                if (documentChanged) {
                    docPart.Document.Save();
                }
            }
        }
    }
}
