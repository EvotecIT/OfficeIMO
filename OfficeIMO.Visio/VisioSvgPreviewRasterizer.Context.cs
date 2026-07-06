using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private sealed class SvgRenderContext {
            private readonly Dictionary<string, XElement> _definitions;
            private readonly HashSet<string> _activeUseIds = new(StringComparer.Ordinal);

            private readonly Func<string, byte[]?>? _imageResolver;

            private SvgRenderContext(SvgStyleSheet styleSheet, Dictionary<string, XElement> definitions, Func<string, byte[]?>? imageResolver, SvgPaintBounds viewportBounds) {
                StyleSheet = styleSheet;
                _definitions = definitions;
                _imageResolver = imageResolver;
                ViewportBounds = viewportBounds;
            }

            internal SvgStyleSheet StyleSheet { get; }

            internal SvgPaintBounds ViewportBounds { get; }

            internal SvgPaintBounds? CurrentPaintBounds { get; private set; }

            internal bool IsVisible { get; private set; } = true;

            internal SvgTextStyle CurrentTextStyle { get; private set; } = SvgTextStyle.Default;

            internal static SvgRenderContext Create(XElement root, SvgPaintBounds viewportBounds, Func<string, byte[]?>? imageResolver = null) =>
                new(SvgStyleSheet.Parse(root), ReadDefinitions(root), imageResolver, viewportBounds);

            internal bool TryGetDefinition(string id, out XElement? definition) =>
                _definitions.TryGetValue(id, out definition);

            internal bool TryEnterUse(string id) => _activeUseIds.Add(id);

            internal void ExitUse(string id) => _activeUseIds.Remove(id);

            internal IDisposable PushPaintBounds(SvgPaintBounds? bounds) {
                SvgPaintBounds? previous = CurrentPaintBounds;
                CurrentPaintBounds = bounds;
                return new PaintBoundsScope(this, previous);
            }

            internal IDisposable PushVisibility(bool? visible) {
                bool previous = IsVisible;
                if (visible.HasValue) {
                    IsVisible = visible.Value;
                }

                return new VisibilityScope(this, previous);
            }

            internal IDisposable PushTextStyle(SvgTextStyle style) {
                SvgTextStyle previous = CurrentTextStyle;
                CurrentTextStyle = style;
                return new TextStyleScope(this, previous);
            }

            internal bool TryGetImageBytes(string href, out byte[]? bytes) {
                bytes = _imageResolver?.Invoke(href);
                return bytes != null && bytes.Length > 0;
            }

            private static Dictionary<string, XElement> ReadDefinitions(XElement root) {
                Dictionary<string, XElement> definitions = new(StringComparer.Ordinal);
                foreach (XElement element in root.Descendants()) {
                    string? id = element.Attribute("id")?.Value;
                    if (!string.IsNullOrWhiteSpace(id) && !definitions.ContainsKey(id!)) {
                        definitions[id!] = element;
                    }
                }

                return definitions;
            }

            private sealed class PaintBoundsScope : IDisposable {
                private readonly SvgRenderContext _context;
                private readonly SvgPaintBounds? _previous;
                private bool _disposed;

                internal PaintBoundsScope(SvgRenderContext context, SvgPaintBounds? previous) {
                    _context = context;
                    _previous = previous;
                }

                public void Dispose() {
                    if (_disposed) {
                        return;
                    }

                    _context.CurrentPaintBounds = _previous;
                    _disposed = true;
                }
            }

            private sealed class VisibilityScope : IDisposable {
                private readonly SvgRenderContext _context;
                private readonly bool _previous;
                private bool _disposed;

                internal VisibilityScope(SvgRenderContext context, bool previous) {
                    _context = context;
                    _previous = previous;
                }

                public void Dispose() {
                    if (_disposed) {
                        return;
                    }

                    _context.IsVisible = _previous;
                    _disposed = true;
                }
            }

            private sealed class TextStyleScope : IDisposable {
                private readonly SvgRenderContext _context;
                private readonly SvgTextStyle _previous;
                private bool _disposed;

                internal TextStyleScope(SvgRenderContext context, SvgTextStyle previous) {
                    _context = context;
                    _previous = previous;
                }

                public void Dispose() {
                    if (_disposed) {
                        return;
                    }

                    _context.CurrentTextStyle = _previous;
                    _disposed = true;
                }
            }
        }

        private readonly struct SvgPaintBounds {
            internal SvgPaintBounds(double left, double top, double width, double height) {
                Left = left;
                Top = top;
                Width = Math.Max(0D, width);
                Height = Math.Max(0D, height);
            }

            internal double Left { get; }

            internal double Top { get; }

            internal double Width { get; }

            internal double Height { get; }

            internal bool HasArea => Width > 0D && Height > 0D;
        }
    }
}
