using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WP14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    public partial class WordShape {
        private OfficeLength? ReadVmlLength(string name) {
            string? text = GetStyleValue(name);
            return string.IsNullOrEmpty(text) ? (OfficeLength?)null : OfficeLength.Parse(text!);
        }

        /// <summary>Gets or sets the displayed shape width in VML or DrawingML. Returns zero when no width is stored.</summary>
        public OfficeLength Width {
            get => _drawing != null ? (TryGetLayoutSnapshot(out var layout) ? layout.Width : OfficeLength.Zero)
                : ReadVmlLength("width") ?? OfficeLength.Zero;
            set => WriteShapeLength(value, horizontal: true);
        }

        /// <summary>Gets or sets the displayed shape height in VML or DrawingML. Returns zero when no height is stored.</summary>
        public OfficeLength Height {
            get => _drawing != null ? (TryGetLayoutSnapshot(out var layout) ? layout.Height : OfficeLength.Zero)
                : ReadVmlLength("height") ?? OfficeLength.Zero;
            set => WriteShapeLength(value, horizontal: false);
        }

        private void WriteShapeLength(OfficeLength value, bool horizontal) {
            if (_drawing == null) {
                SetStyleValue(horizontal ? "width" : "height", FormatVmlLength(value));
                return;
            }
            long emus = value.ToEmus();
            if (emus <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Shape dimensions must be positive.");
            WordDrawingLayoutReader.SetExtent(_drawing, horizontal ? emus : (long?)null, horizontal ? null : (long?)emus);
            var properties = _wpsShape?.GetFirstChild<Wps.ShapeProperties>();
            if (properties == null) return;
            var transform = properties.GetFirstChild<A.Transform2D>();
            if (transform == null) {
                transform = new A.Transform2D();
                properties.AddChild(transform, true);
            }
            if (transform.Extents == null) {
                var frame = _drawing.Inline?.Extent ?? _drawing.Anchor?.Extent;
                transform.Extents = new A.Extents { Cx = frame?.Cx?.Value ?? 0L, Cy = frame?.Cy?.Value ?? 0L };
            }
            if (horizontal) transform.Extents.Cx = emus;
            else transform.Extents.Cy = emus;
        }

        /// <summary>Gets or sets the horizontal shape offset. Inline DrawingML shapes return null and become anchored when an offset is set. Null clears the explicit offset.</summary>
        public OfficeLength? Left {
            get => _drawing != null ? (TryGetLayoutSnapshot(out var layout) ? layout.HorizontalOffset : null) : ReadVmlLength("margin-left");
            set {
                if (_drawing != null) {
                    WriteDrawingOffset(value, horizontal: true);
                    return;
                }
                if (value == null) {
                    RemoveStyleValue("margin-left");
                } else {
                    SetStyleValue("margin-left", FormatVmlLength(value.Value));
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        /// <summary>Gets or sets the vertical shape offset. Inline DrawingML shapes return null and become anchored when an offset is set. Null clears the explicit offset.</summary>
        public OfficeLength? Top {
            get => _drawing != null ? (TryGetLayoutSnapshot(out var layout) ? layout.VerticalOffset : null) : ReadVmlLength("margin-top");
            set {
                if (_drawing != null) {
                    WriteDrawingOffset(value, horizontal: false);
                    return;
                }
                if (value == null) {
                    RemoveStyleValue("margin-top");
                } else {
                    SetStyleValue("margin-top", FormatVmlLength(value.Value));
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        private void WriteDrawingOffset(OfficeLength? value, bool horizontal) {
            // wp:posOffset is signed 32-bit EMUs; reject overflow before converting an inline or replacing alignment.
            int? emus = value.HasValue ? checked((int)value.Value.ToEmus()) : (int?)null;
            var anchor = _drawing!.Anchor;
            if (anchor == null && _drawing.Inline is DW.Inline inline) {
                if (!emus.HasValue) return;
                anchor = WordTextBox.ConvertInlineToAnchor(inline, WordImageTextWrapping.Square);
                anchor.AnchorId = inline.AnchorId;
                anchor.EditId = inline.EditId;
                anchor.DistanceFromTop = inline.DistanceFromTop?.Value ?? 0U;
                anchor.DistanceFromBottom = inline.DistanceFromBottom?.Value ?? 0U;
                anchor.DistanceFromLeft = inline.DistanceFromLeft?.Value ?? 0U;
                anchor.DistanceFromRight = inline.DistanceFromRight?.Value ?? 0U;
                anchor.GetFirstChild<WP14.RelativeWidth>()?.Remove();
                anchor.GetFirstChild<WP14.RelativeHeight>()?.Remove();
                anchor.HorizontalPosition = new DW.HorizontalPosition {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page,
                    HorizontalAlignment = new DW.HorizontalAlignment { Text = "left" }
                };
                anchor.VerticalPosition = new DW.VerticalPosition {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page,
                    VerticalAlignment = new DW.VerticalAlignment { Text = "top" }
                };
                // Keep the source graphic and its model references; the image conversion clones its payload.
                if (inline.Graphic is A.Graphic graphic) {
                    graphic.Remove();
                    anchor.ReplaceChild(graphic, anchor.GetFirstChild<A.Graphic>()!);
                }
                _drawing.ReplaceChild(anchor, inline);
            }
            if (anchor == null) return;
            if (anchor.SimplePos?.Value == true) {
                if (emus.HasValue && anchor.SimplePosition != null) {
                    if (horizontal) anchor.SimplePosition.X = emus.Value;
                    else anchor.SimplePosition.Y = emus.Value;
                    return;
                }
                // Clearing one simple coordinate switches to explicit positioning while preserving the other axis.
                int other = checked((int)(horizontal ? anchor.SimplePosition?.Y?.Value ?? 0L : anchor.SimplePosition?.X?.Value ?? 0L));
                anchor.HorizontalPosition = new DW.HorizontalPosition {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page,
                    PositionOffset = new DW.PositionOffset((horizontal ? 0 : other).ToString(CultureInfo.InvariantCulture))
                };
                anchor.VerticalPosition = new DW.VerticalPosition {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page,
                    PositionOffset = new DW.PositionOffset((horizontal ? other : 0).ToString(CultureInfo.InvariantCulture))
                };
                anchor.SimplePos = false;
            }
            if (horizontal) {
                var position = anchor.HorizontalPosition;
                if (position == null) {
                    if (!emus.HasValue) return;
                    position = new DW.HorizontalPosition { RelativeFrom = DW.HorizontalRelativePositionValues.Page };
                    anchor.HorizontalPosition = position;
                }
                position.PositionOffset = emus.HasValue ? new DW.PositionOffset(emus.Value.ToString(CultureInfo.InvariantCulture)) : null;
                if (emus.HasValue) position.HorizontalAlignment?.Remove();
                else position.HorizontalAlignment ??= new DW.HorizontalAlignment { Text = "left" };
            } else {
                var position = anchor.VerticalPosition;
                if (position == null) {
                    if (!emus.HasValue) return;
                    position = new DW.VerticalPosition { RelativeFrom = DW.VerticalRelativePositionValues.Page };
                    anchor.VerticalPosition = position;
                }
                position.PositionOffset = emus.HasValue ? new DW.PositionOffset(emus.Value.ToString(CultureInfo.InvariantCulture)) : null;
                if (emus.HasValue) position.VerticalAlignment?.Remove();
                else position.VerticalAlignment ??= new DW.VerticalAlignment { Text = "top" };
            }
        }

        /// <summary>Gets or sets the outline thickness. Null removes the explicit thickness while preserving other outline formatting.</summary>
        public OfficeLength? StrokeWeight {
            get {
                string? text = _rectangle?.StrokeWeight?.Value ?? _roundRectangle?.StrokeWeight?.Value ??
                    _ellipse?.StrokeWeight?.Value ?? _polygon?.StrokeWeight?.Value ?? _line?.StrokeWeight?.Value ?? _shape?.StrokeWeight?.Value;
                if (!string.IsNullOrEmpty(text)) return OfficeLength.Parse(text!);
                var outline = _wpsShape?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties>()?
                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();
                return outline?.Width?.Value is int emus ? OfficeLength.Emus(emus) : (OfficeLength?)null;
            }
            set {
                // Validate the native DrawingML boundary before changing either representation.
                int? emus = _wpsShape != null && value.HasValue ? checked((int)value.Value.ToEmus()) : (int?)null;
                string? text = value.HasValue ? FormatVmlLength(value.Value) : null;
                if (_rectangle != null) _rectangle.StrokeWeight = text;
                if (_roundRectangle != null) _roundRectangle.StrokeWeight = text;
                if (_ellipse != null) _ellipse.StrokeWeight = text;
                if (_polygon != null) _polygon.StrokeWeight = text;
                if (_line != null) _line.StrokeWeight = text;
                if (_shape != null) _shape.StrokeWeight = text;
                if (_wpsShape != null) {
                    var properties = _wpsShape!.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties>();
                    if (properties != null) {
                        var outline = properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();
                        if (outline == null && emus.HasValue) {
                            outline = new DocumentFormat.OpenXml.Drawing.Outline();
                            properties.Append(outline);
                        }
                        if (outline != null) {
                            outline.Width = emus.HasValue ? new DocumentFormat.OpenXml.Int32Value(emus.Value) : null;
                        }
                    }
                }
            }
        }
    }
}
