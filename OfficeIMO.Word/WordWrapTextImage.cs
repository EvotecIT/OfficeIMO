using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word {
    public enum WrapTextImage {
        InLineWithText,
        Square,
        Tight,
        Through,
        TopAndBottom,
        BehindText,
        InFrontOfText,
    }

    public class WordWrapTextImage {
        private WordWrapTextImage(WrapTextImage wrapTextImage) {

        }

        public static Anchor AppendWrapTextImage(Anchor anchor, WrapTextImage wrapImage) {
            if (wrapImage == WrapTextImage.Square) {
                WrapSquare wrapSquare1 = new WrapSquare() {
                    WrapText = WrapTextValues.BothSides
                };
                anchor.Append(wrapSquare1);
            } else if (wrapImage == WrapTextImage.Tight) {
                WrapTight wrapTight1 = WordWrapTextImage.WrapTight;
                anchor.Append(wrapTight1);
            } else if (wrapImage == WrapTextImage.Through) {
                WrapThrough wrapThrough1 = WordWrapTextImage.WrapThrough;
                anchor.Append(wrapThrough1);
            } else if (wrapImage == WrapTextImage.TopAndBottom) {
                WrapTopBottom wrapTopBottom1 = WordWrapTextImage.WrapTopBottom;
                anchor.Append(wrapTopBottom1);
            } else if (wrapImage == WrapTextImage.BehindText) {
                WrapNone wrapNone1 = new WrapNone();
                anchor.Append(wrapNone1);
                anchor.BehindDoc = true;
            } else if (wrapImage == WrapTextImage.InFrontOfText) {
                WrapNone wrapNone1 = new WrapNone();
                anchor.Append(wrapNone1);
                anchor.BehindDoc = false;
            } else {
                throw new InvalidOperationException("WrapTextImage: " + wrapImage + " not supported yet.");
            }
            return anchor;
        }

        public static WrapTextImage? GetWrapTextImage(Anchor anchor, Inline inline) {
            if (anchor != null) {
                var wrapSquare = anchor.OfType<WrapSquare>().FirstOrDefault();
                if (wrapSquare != null) {
                    return WrapTextImage.Square;
                }
                var wrapTight = anchor.OfType<WrapTight>().FirstOrDefault();
                if (wrapTight != null) {
                    return WrapTextImage.Tight;
                }
                var wrapThrough = anchor.OfType<WrapThrough>().FirstOrDefault();
                if (wrapThrough != null) {
                    return WrapTextImage.Through;
                }
                var wrapTopAndBottom = anchor.OfType<WrapTopBottom>().FirstOrDefault();
                if (wrapTopAndBottom != null) {
                    return WrapTextImage.TopAndBottom;
                }
                var wrapNone = anchor.OfType<WrapNone>().FirstOrDefault();
                var behindDoc = anchor.BehindDoc;
                if (wrapNone != null && behindDoc != null && behindDoc.Value == true) {
                    return WrapTextImage.BehindText;
                } else if (wrapNone != null && behindDoc != null && behindDoc.Value == false) {
                    return WrapTextImage.InFrontOfText;
                }
            } else if (inline != null) {
                return WrapTextImage.InLineWithText;
            }
            return null;
        }

        public static void SetWrapTextImage(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, Anchor anchor, Inline inline, WrapTextImage? wrapImage) {
            var currentWrap = GetWrapTextImage(anchor, inline);
            if (currentWrap == wrapImage) {
                // nothing to do
                return;
            }
            if (anchor != null) {
                if (wrapImage == WrapTextImage.InLineWithText) {
                    var convertedInline = WordTextBox.ConvertAnchorToInline(anchor);
                    drawing.Append(convertedInline);
                    drawing.OfType<Anchor>().FirstOrDefault()?.Remove();
                } else {
                    // remove current Wrap
                    if (currentWrap == WrapTextImage.Square) {
                        anchor.OfType<WrapSquare>().FirstOrDefault()?.Remove();
                    } else if (currentWrap == WrapTextImage.Tight) {
                        anchor.OfType<WrapTight>().FirstOrDefault()?.Remove();
                    } else if (currentWrap == WrapTextImage.Through) {
                        anchor.OfType<WrapThrough>().FirstOrDefault()?.Remove();
                    } else if (currentWrap == WrapTextImage.TopAndBottom) {
                        anchor.OfType<WrapTopBottom>().FirstOrDefault()?.Remove();
                    } else if (currentWrap == WrapTextImage.BehindText) {
                        anchor.OfType<WrapNone>().FirstOrDefault()?.Remove();
                        anchor.BehindDoc = true;
                    } else if (currentWrap == WrapTextImage.InFrontOfText) {
                        anchor.OfType<WrapNone>().FirstOrDefault()?.Remove();
                        anchor.BehindDoc = false;
                    } else if (currentWrap == WrapTextImage.InLineWithText) {
                        // this won't really happen
                    }

                    // wrap needs to be inserted after extent or it will not work
                    var extent = anchor.Elements<Extent>().FirstOrDefault();
                    if (extent != null) {
                        if (wrapImage == WrapTextImage.Square) {
                            var wrap = new WrapSquare() { WrapText = WrapTextValues.BothSides };
                            extent.InsertAfterSelf(wrap);
                        } else if (wrapImage == WrapTextImage.Tight) {
                            anchor.Append(new WrapTight() { WrapText = WrapTextValues.BothSides });
                        } else if (wrapImage == WrapTextImage.Through) {
                            var wrap = WordWrapTextImage.WrapThrough;
                            extent.InsertAfterSelf(wrap);
                        } else if (wrapImage == WrapTextImage.TopAndBottom) {
                            var wrap = WordWrapTextImage.WrapTopBottom;
                            extent.InsertAfterSelf(wrap);
                        } else if (wrapImage == WrapTextImage.BehindText) {
                            var wrap = new WrapNone();
                            extent.InsertAfterSelf(wrap);
                            anchor.BehindDoc = false;
                        } else if (wrapImage == WrapTextImage.InFrontOfText) {
                            var wrap = new WrapNone();
                            extent.InsertAfterSelf(wrap);
                            anchor.BehindDoc = true;
                        } else if (wrapImage == WrapTextImage.InLineWithText) {
                            throw new InvalidOperationException("WrapTextImage.InLineWithText should be handled before.");
                        }
                    } else {
                        throw new InvalidOperationException("Extent is missing. Weird. Shouldn't happen.");
                    }
                }
            } else if (inline != null) {
                if (wrapImage == WrapTextImage.InLineWithText) {
                    // nothing to do
                    return;
                } else {
                    var convertedAnchor = WordTextBox.ConvertInlineToAnchor(inline, wrapImage.Value);
                    drawing.Append(convertedAnchor);
                    drawing.OfType<Inline>().FirstOrDefault()?.Remove();
                }
            }
        }


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
