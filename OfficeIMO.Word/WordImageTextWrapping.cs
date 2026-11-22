using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Specifies the available options for how text should wrap around an image.
    /// </summary>
    public enum WordImageTextWrapping {
        /// <summary>No wrapping is applied and the image is placed inline with the surrounding text.</summary>
        InLineWithText,
        /// <summary>Wrap text around the image using a square boundary.</summary>
        Square,
        /// <summary>Wrap text tightly around the outline of the image.</summary>
        Tight,
        /// <summary>Allow text to flow through transparent regions of the image.</summary>
        Through,
        /// <summary>Place text above and below the image leaving the sides clear.</summary>
        TopAndBottom,
        /// <summary>Position the image behind the text.</summary>
        BehindText,
        /// <summary>Position the image in front of the text.</summary>
        InFrontOfText,
    }

    /// <summary>
    /// Provides helper methods for configuring image wrap options on Wordprocessing elements.
    /// </summary>
    internal class WordWrapTextImage {
        private WordWrapTextImage(WordImageTextWrapping wrapTextImage) {

        }

        /// <summary>
        /// Appends the appropriate wrapping element for the specified option to the given anchor.
        /// </summary>
        /// <param name="anchor">The anchor to which the wrapping should be applied.</param>
        /// <param name="wrapImage">The desired wrapping option.</param>
        /// <returns>The modified anchor.</returns>
        public static Anchor AppendWrapTextImage(Anchor anchor, WordImageTextWrapping wrapImage) {
            anchor.Append(CreateWrapElement(wrapImage));
            anchor.BehindDoc = wrapImage == WordImageTextWrapping.BehindText;
            return anchor;
        }

        /// <summary>
        /// Creates the anchor wrap element for a floating wrap option.
        /// </summary>
        private static OpenXmlElement CreateWrapElement(WordImageTextWrapping wrapImage) => wrapImage switch {
            WordImageTextWrapping.Square => new WrapSquare() { WrapText = WrapTextValues.BothSides },
            WordImageTextWrapping.Tight => WrapTight,
            WordImageTextWrapping.Through => WrapThrough,
            WordImageTextWrapping.TopAndBottom => WrapTopBottom,
            WordImageTextWrapping.BehindText or WordImageTextWrapping.InFrontOfText => new WrapNone(),
            _ => throw new InvalidOperationException("WrapTextImage: " + wrapImage + " not supported yet.")
        };

        /// <summary>
        /// Returns the currently applied wrapping option for the specified drawing.
        /// </summary>
        /// <param name="anchor">Anchor element associated with the drawing.</param>
        /// <param name="inline">Inline element associated with the drawing.</param>
        /// <returns>The wrap option or <c>null</c> if none can be determined.</returns>
        public static WordImageTextWrapping? GetWrapTextImage(Anchor anchor, Inline inline) {
            if (anchor != null) {
                var wrapSquare = anchor.OfType<WrapSquare>().FirstOrDefault();
                if (wrapSquare != null) {
                    return WordImageTextWrapping.Square;
                }
                var wrapTight = anchor.OfType<WrapTight>().FirstOrDefault();
                if (wrapTight != null) {
                    return WordImageTextWrapping.Tight;
                }
                var wrapThrough = anchor.OfType<WrapThrough>().FirstOrDefault();
                if (wrapThrough != null) {
                    return WordImageTextWrapping.Through;
                }
                var wrapTopAndBottom = anchor.OfType<WrapTopBottom>().FirstOrDefault();
                if (wrapTopAndBottom != null) {
                    return WordImageTextWrapping.TopAndBottom;
                }
                var wrapNone = anchor.OfType<WrapNone>().FirstOrDefault();
                var behindDoc = anchor.BehindDoc;
                if (wrapNone != null && behindDoc != null && behindDoc.Value == true) {
                    return WordImageTextWrapping.BehindText;
                } else if (wrapNone != null && behindDoc != null && behindDoc.Value == false) {
                    return WordImageTextWrapping.InFrontOfText;
                }
            } else if (inline != null) {
                return WordImageTextWrapping.InLineWithText;
            }
            return null;
        }

        /// <summary>
        /// Sets the wrapping option for the provided drawing.
        /// </summary>
        /// <param name="drawing">The WordDrawing element to update.</param>
        /// <param name="anchor">Anchor element associated with the drawing.</param>
        /// <param name="inline">Inline element associated with the drawing.</param>
        /// <param name="wrapImage">The desired wrapping option.</param>
        public static void SetWrapTextImage(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, Anchor anchor, Inline inline, WordImageTextWrapping? wrapImage) {
            var currentWrap = GetWrapTextImage(anchor, inline);
            ArgumentNullException.ThrowIfNull(wrapImage);
            if (currentWrap == wrapImage) {
                // nothing to do
                return;
            }
            if (anchor != null) {
                if (wrapImage == WordImageTextWrapping.InLineWithText) {
                    var convertedInline = WordTextBox.ConvertAnchorToInline(anchor);
                    drawing.Append(convertedInline);
                    drawing.OfType<Anchor>().FirstOrDefault()?.Remove();
                } else {
                    foreach (OpenXmlElement existing in anchor.ChildElements
                        .Where(static element => element is WrapNone or WrapSquare or DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTight
                            or DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapThrough or DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTopBottom).ToList()) {
                        existing.Remove();
                    }
                    // The wrap follows the extent and effect extent in the anchor's schema order.
                    OpenXmlElement? previous = (OpenXmlElement?)anchor.GetFirstChild<EffectExtent>() ?? anchor.GetFirstChild<Extent>();
                    if (previous == null) throw new InvalidOperationException("Extent is missing. Weird. Shouldn't happen.");
                    previous.InsertAfterSelf(CreateWrapElement(wrapImage.Value));
                    // Word saves behindDoc="1" only for behind-text pictures; a wrapped picture paints by z-order.
                    anchor.BehindDoc = wrapImage == WordImageTextWrapping.BehindText;
                }
            } else if (inline != null) {
                if (wrapImage == WordImageTextWrapping.InLineWithText) {
                    // nothing to do
                    return;
                } else {
                    var convertedAnchor = WordTextBox.ConvertInlineToAnchor(inline, wrapImage.Value);
                    drawing.Append(convertedAnchor);
                    drawing.OfType<Inline>().FirstOrDefault()?.Remove();
                }
            }
        }


        /// <summary>
        /// Gets a <see cref="WrapTopBottom"/> instance used for top and bottom wrapping.
        /// </summary>
        public static WrapTopBottom WrapTopBottom {
            get {
                WrapTopBottom wrapTopBottom1 = new WrapTopBottom() {
                    //Values don't seem to matter
                    //DistanceFromTop = (UInt32Value)20U,
                    //DistanceFromBottom = (UInt32Value)20U
                };
                return wrapTopBottom1;
            }
        }

        /// <summary>
        /// Gets a <see cref="WrapThrough"/> instance used for through wrapping.
        /// </summary>
        public static WrapThrough WrapThrough {
            get {
                WrapThrough wrapThrough1 = new WrapThrough() { WrapText = WrapTextValues.BothSides };
                WrapPolygon wrapPolygon1 = new WrapPolygon() { Edited = false };
                StartPoint startPoint1 = new StartPoint() { X = 0L, Y = 0L };
                // the values are probably wrong and content oriented
                // would require some more research on how to calculate them
                var lineTo1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 0L, Y = 21384L };
                var lineTo2 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 21384L, Y = 21384L };
                var lineTo3 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 21384L, Y = 0L };
                var lineTo4 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 0L, Y = 0L };

                wrapPolygon1.Append(startPoint1);
                wrapPolygon1.Append(lineTo1);
                wrapPolygon1.Append(lineTo2);
                wrapPolygon1.Append(lineTo3);
                wrapPolygon1.Append(lineTo4);
                wrapThrough1.Append(wrapPolygon1);
                return wrapThrough1;
            }
        }

        /// <summary>
        /// Gets a <see cref="WrapTight"/> instance used for tight wrapping.
        /// </summary>
        public static WrapTight WrapTight {
            get {
                WrapTight wrapTight1 = new WrapTight() { WrapText = WrapTextValues.BothSides };
                WrapPolygon wrapPolygon1 = new WrapPolygon() { Edited = false };
                StartPoint startPoint1 = new StartPoint() { X = 0L, Y = 0L };
                // the values are probably wrong and content oriented
                // would require some more research on how to calculate them
                var lineTo1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 0L, Y = 21384L };
                var lineTo2 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 21384L, Y = 21384L };
                var lineTo3 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 21384L, Y = 0L };
                var lineTo4 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo() { X = 0L, Y = 0L };

                wrapPolygon1.Append(startPoint1);
                wrapPolygon1.Append(lineTo1);
                wrapPolygon1.Append(lineTo2);
                wrapPolygon1.Append(lineTo3);
                wrapPolygon1.Append(lineTo4);

                wrapTight1.Append(wrapPolygon1);
                return wrapTight1;
            }
        }
    }
}
