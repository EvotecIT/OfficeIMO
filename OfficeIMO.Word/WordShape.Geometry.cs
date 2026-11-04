namespace OfficeIMO.Word {
    public partial class WordShape {
        private OfficeLength? ReadVmlLength(string name) {
            string? text = GetStyleValue(name);
            return string.IsNullOrEmpty(text) ? (OfficeLength?)null : OfficeLength.Parse(text!);
        }

        /// <summary>Gets or sets the VML shape width. Returns zero when no VML width is stored.</summary>
        public OfficeLength Width {
            get => ReadVmlLength("width") ?? OfficeLength.Zero;
            set => SetStyleValue("width", FormatVmlLength(value));
        }

        /// <summary>Gets or sets the VML shape height. Returns zero when no VML height is stored.</summary>
        public OfficeLength Height {
            get => ReadVmlLength("height") ?? OfficeLength.Zero;
            set => SetStyleValue("height", FormatVmlLength(value));
        }

        /// <summary>Gets or sets the VML left offset. Null removes the explicit offset.</summary>
        public OfficeLength? Left {
            get => ReadVmlLength("margin-left");
            set {
                if (value == null) {
                    RemoveStyleValue("margin-left");
                } else {
                    SetStyleValue("margin-left", FormatVmlLength(value.Value));
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        /// <summary>Gets or sets the VML top offset. Null removes the explicit offset.</summary>
        public OfficeLength? Top {
            get => ReadVmlLength("margin-top");
            set {
                if (value == null) {
                    RemoveStyleValue("margin-top");
                } else {
                    SetStyleValue("margin-top", FormatVmlLength(value.Value));
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        /// <summary>Gets or sets the outline thickness. Null preserves the existing format-specific clearing behavior.</summary>
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
                // Preserve the existing DrawingML null no-op; VML null clears its attribute.
                if (emus.HasValue) {
                    var properties = _wpsShape!.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties>();
                    if (properties != null) {
                        var outline = properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();
                        if (outline == null) {
                            outline = new DocumentFormat.OpenXml.Drawing.Outline();
                            properties.Append(outline);
                        }
                        outline.Width = emus.Value;
                    }
                }
            }
        }
    }
}
