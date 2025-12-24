using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a formatted text run within a paragraph.
    /// </summary>
    public class PowerPointTextRun {
        internal PowerPointTextRun(A.Run run) {
            Run = run;
        }

        internal A.Run Run { get; }

        /// <summary>
        /// Text content of the run.
        /// </summary>
        public string Text {
            get => Run.Text?.Text ?? string.Empty;
            set {
                Run.Text ??= new A.Text();
                Run.Text.Text = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is bold.
        /// </summary>
        public bool Bold {
            get => Run.RunProperties?.Bold?.Value == true;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Bold = value ? true : null;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is italic.
        /// </summary>
        public bool Italic {
            get => Run.RunProperties?.Italic?.Value == true;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Italic = value ? true : null;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is underlined.
        /// </summary>
        public bool Underline {
            get => Run.RunProperties?.Underline?.Value == A.TextUnderlineValues.Single;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Underline = value ? A.TextUnderlineValues.Single : null;
            }
        }

        /// <summary>
        /// Gets or sets the font size in points.
        /// </summary>
        public int? FontSize {
            get {
                int? size = Run.RunProperties?.FontSize?.Value;
                return size != null ? size / 100 : null;
            }
            set {
                A.RunProperties props = EnsureRunProperties();
                props.FontSize = value != null ? value * 100 : null;
            }
        }

        /// <summary>
        /// Gets or sets the font name (Latin).
        /// </summary>
        public string? FontName {
            get => Run.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.RemoveAllChildren<A.LatinFont>();
                if (value != null) {
                    props.Append(new A.LatinFont { Typeface = value });
                }
            }
        }

        /// <summary>
        /// Gets or sets the text color in hexadecimal format (e.g. "FF0000").
        /// </summary>
        public string? Color {
            get => Run.RunProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
            set {
                A.RunProperties props = EnsureRunProperties();
                var latin = props.GetFirstChild<A.LatinFont>();
                var ea = props.GetFirstChild<A.EastAsianFont>();
                var cs = props.GetFirstChild<A.ComplexScriptFont>();

                props.RemoveAllChildren<A.SolidFill>();
                props.RemoveAllChildren<A.LatinFont>();
                props.RemoveAllChildren<A.EastAsianFont>();
                props.RemoveAllChildren<A.ComplexScriptFont>();

                if (value != null) {
                    props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                }

                if (latin != null) props.Append((A.LatinFont)latin.CloneNode(true));
                if (ea != null) props.Append((A.EastAsianFont)ea.CloneNode(true));
                if (cs != null) props.Append((A.ComplexScriptFont)cs.CloneNode(true));
            }
        }

        private A.RunProperties EnsureRunProperties() {
            return Run.RunProperties ??= new A.RunProperties();
        }
    }
}
