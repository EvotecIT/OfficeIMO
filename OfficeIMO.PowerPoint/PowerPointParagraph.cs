using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a paragraph within a textbox.
    /// </summary>
    public class PowerPointParagraph {
        internal PowerPointParagraph(A.Paragraph paragraph) {
            Paragraph = paragraph;
        }

        internal A.Paragraph Paragraph { get; }

        /// <summary>
        /// Text content of the paragraph.
        /// </summary>
        public string Text {
            get => Paragraph.InnerText ?? string.Empty;
            set {
                A.EndParagraphRunProperties? endProps = Paragraph.GetFirstChild<A.EndParagraphRunProperties>();
                endProps?.Remove();
                Paragraph.RemoveAllChildren<A.Run>();
                A.Run run = new(new A.Text(value ?? string.Empty));
                Paragraph.Append(run);
                if (endProps != null) {
                    Paragraph.Append(endProps);
                }
            }
        }

        /// <summary>
        /// Runs within the paragraph.
        /// </summary>
        public IReadOnlyList<PowerPointTextRun> Runs =>
            Paragraph.Elements<A.Run>().Select(r => new PowerPointTextRun(r)).ToList();

        /// <summary>
        /// Adds a run to the paragraph.
        /// </summary>
        public PowerPointTextRun AddRun(string text, Action<PowerPointTextRun>? configure = null) {
            A.Run run = new(new A.Text(text ?? string.Empty));
            A.EndParagraphRunProperties? endProps = Paragraph.GetFirstChild<A.EndParagraphRunProperties>();
            if (endProps != null) {
                Paragraph.InsertBefore(run, endProps);
            } else {
                Paragraph.Append(run);
            }
            var wrapper = new PowerPointTextRun(run);
            configure?.Invoke(wrapper);
            return wrapper;
        }

        /// <summary>
        /// Gets or sets paragraph alignment.
        /// </summary>
        public A.TextAlignmentTypeValues? Alignment {
            get => Paragraph.ParagraphProperties?.Alignment?.Value;
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Alignment = value;
            }
        }

        /// <summary>
        /// Gets or sets the bullet/list level (0-8).
        /// </summary>
        public int? Level {
            get => Paragraph.ParagraphProperties?.Level?.Value;
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Level = value;
            }
        }

        /// <summary>
        /// Gets or sets paragraph indentation in points.
        /// </summary>
        public double? IndentPoints {
            get => FromTextCoordinate(Paragraph.ParagraphProperties?.Indent?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Indent = ToTextCoordinate(value);
            }
        }

        /// <summary>
        /// Gets or sets line spacing in points.
        /// </summary>
        public double? LineSpacingPoints {
            get => FromSpacingPoints(Paragraph.ParagraphProperties?.LineSpacing?.SpacingPoints?.Val?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                if (value == null) {
                    props.LineSpacing = null;
                    return;
                }
                props.LineSpacing = new A.LineSpacing(new A.SpacingPoints { Val = ToSpacingPoints(value.Value) });
            }
        }

        /// <summary>
        /// Gets or sets space before the paragraph in points.
        /// </summary>
        public double? SpaceBeforePoints {
            get => FromSpacingPoints(Paragraph.ParagraphProperties?.SpaceBefore?.SpacingPoints?.Val?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                if (value == null) {
                    props.SpaceBefore = null;
                    return;
                }
                props.SpaceBefore = new A.SpaceBefore(new A.SpacingPoints { Val = ToSpacingPoints(value.Value) });
            }
        }

        /// <summary>
        /// Gets or sets space after the paragraph in points.
        /// </summary>
        public double? SpaceAfterPoints {
            get => FromSpacingPoints(Paragraph.ParagraphProperties?.SpaceAfter?.SpacingPoints?.Val?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                if (value == null) {
                    props.SpaceAfter = null;
                    return;
                }
                props.SpaceAfter = new A.SpaceAfter(new A.SpacingPoints { Val = ToSpacingPoints(value.Value) });
            }
        }

        /// <summary>
        /// Applies a character bullet to the paragraph.
        /// </summary>
        public void SetBullet(char bulletChar = '\u2022') {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            props.Append(new A.CharacterBullet { Char = bulletChar.ToString() });
        }

        /// <summary>
        /// Applies an auto-numbered bullet to the paragraph.
        /// </summary>
        public void SetNumbered(A.TextAutoNumberSchemeValues style, int startAt = 1) {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            props.Append(new A.AutoNumberedBullet { Type = style, StartAt = startAt });
        }

        /// <summary>
        /// Applies a default auto-numbered bullet (Arabic period) to the paragraph.
        /// </summary>
        public void SetNumbered(int startAt = 1) {
            SetNumbered(A.TextAutoNumberSchemeValues.ArabicPeriod, startAt);
        }

        /// <summary>
        /// Clears any bullet/numbering from the paragraph.
        /// </summary>
        public void ClearBullet() {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            props.Append(new A.NoBullet());
        }

        private static void ClearBulletInternal(A.ParagraphProperties props) {
            props.RemoveAllChildren<A.BulletFont>();
            props.RemoveAllChildren<A.CharacterBullet>();
            props.RemoveAllChildren<A.AutoNumberedBullet>();
            props.RemoveAllChildren<A.NoBullet>();
        }

        private A.ParagraphProperties EnsureParagraphProperties() {
            return Paragraph.ParagraphProperties ??= new A.ParagraphProperties();
        }

        private static int? ToTextCoordinate(double? points) {
            return points != null ? (int)Math.Round(points.Value * 100) : null;
        }

        private static double? FromTextCoordinate(int? value) {
            return value != null ? value.Value / 100d : null;
        }

        private static int ToSpacingPoints(double points) {
            return (int)Math.Round(points * 100);
        }

        private static double? FromSpacingPoints(int? value) {
            return value != null ? value.Value / 100d : null;
        }
    }
}
