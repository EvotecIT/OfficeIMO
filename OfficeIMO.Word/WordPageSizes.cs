using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum WordPageSize {
        Unknown,
        Letter,
        Legal,
        Statement,
        Executive,
        A3,
        A4,
        A5,
        A6,
        B5
    }

    public class WordPageSizes {
        private readonly WordSection _section;
        private readonly WordDocument _document;

        public WordPageSize? PageSize {
            get {
                var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                if (pageSize != null) {
                    foreach (WordPageSize wordPageSize in Enum.GetValues(typeof(WordPageSize))) {
                        if (wordPageSize == WordPageSize.Unknown) {
                            continue;
                        }

                        var pageSizeBuiltin = GetDefault(wordPageSize);
                        if ((pageSizeBuiltin.Width == null && pageSize.Height == null) &&
                            (pageSizeBuiltin.Height == null && pageSize.Height == null) &&
                            (pageSizeBuiltin.Code == null && pageSize.Code == null)) {
                            return wordPageSize;
                        }

                        if (pageSizeBuiltin.Width != null && pageSize.Width != null &&
                            pageSizeBuiltin.Height != null && pageSize.Height != null &&
                            pageSizeBuiltin.Code != null && pageSize.Code != null &&
                            pageSizeBuiltin.Width == pageSize.Width &&
                            pageSizeBuiltin.Height == pageSize.Height &&
                            pageSizeBuiltin.Code == pageSize.Code) {
                            return wordPageSize;
                        }
                    }
                    // page size is set but we don't know what it is
                    return WordPageSize.Unknown;
                } else {
                    // not set page size
                    return null;
                }
            }
            set => SetPageSize(value);
        }

        private void SetPageSize(WordPageSize? wordPageSize) {
            if (wordPageSize == null) {
                var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                if (pageSize != null) {
                    pageSize.Remove();
                }
            } else {
                var pageSizeSettings = GetDefault(wordPageSize);
                if (pageSizeSettings == null) {
                    var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                    if (pageSize != null) {
                        pageSize.Remove();
                    }
                } else {
                    var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();
                    if (pageSize == null) {
                        _section._sectionProperties.Append(pageSizeSettings);
                    } else {
                        pageSize.Remove();
                        _section._sectionProperties.Append(pageSizeSettings);
                    }
                }
            }
        }

        public UInt32Value Width {
            get {
                var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                if (pageSize != null) {
                    return pageSize.Width.Value;
                }

                return null;
            }
            set {
                if (_section._sectionProperties != null) {
                    var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                    if (pageSize == null) {
                        pageSize = new PageSize();
                    }

                    pageSize.Width = value;
                }
            }
        }

        public UInt32Value Height {
            get {
                var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                if (pageSize != null) {
                    return pageSize.Height.Value;
                }

                return null;
            }
            set {
                if (_section._sectionProperties != null) {
                    var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                    if (pageSize == null) {
                        pageSize = new PageSize();
                    }

                    pageSize.Height = value;
                }
            }
        }

        public UInt16Value Code {
            get {
                var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                if (pageSize != null) {
                    return pageSize.Code;
                }

                return null;
            }
            set {
                if (_section._sectionProperties != null) {
                    var pageSize = _section._sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
                    if (pageSize == null) {
                        pageSize = new PageSize();
                    }

                    pageSize.Code = value;
                }
            }
        }

        internal static PageOrientationValues GetOrientation(SectionProperties sectionProperties) {
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            if (pageSize == null) {
                return PageOrientationValues.Portrait;
            }

            if (pageSize.Orient != null) {
                return pageSize.Orient.Value;
            }

            return PageOrientationValues.Portrait;
        }

        internal static void SetOrientation(SectionProperties sectionProperties, PageOrientationValues pageOrientationValue) {
            var pageSize = sectionProperties.Descendants<PageSize>().FirstOrDefault();
            if (pageSize == null) {
                // we need to setup default values for A4 
                pageSize = WordPageSizes.A4;
                pageSize.Orient = PageOrientationValues.Portrait;
                sectionProperties.Append(pageSize);
            }
            if (pageSize.Orient == null) {
                pageSize.Orient = PageOrientationValues.Portrait;
            }
            if (pageSize.Orient != pageOrientationValue) {
                // changing orientation is not enough, we need to change width with height and vice versa
                var width = pageSize.Width;
                var height = pageSize.Height;
                pageSize.Width = height;
                pageSize.Height = width;

                pageSize.Orient = pageOrientationValue;
            }
        }

        public PageOrientationValues Orientation {
            get => GetOrientation(_section._sectionProperties);
            set => SetOrientation(_section._sectionProperties, value);
        }

        public WordPageSizes(WordDocument wordDocument, WordSection wordSection) {
            _section = wordSection;
            _document = wordDocument;
        }

        private static PageSize GetDefault(WordPageSize? pageSize) {
            switch (pageSize) {
                case WordPageSize.A3: return A3;
                case WordPageSize.A4: return A4;
                case WordPageSize.A5: return A5;
                case WordPageSize.Executive: return Executive;
                case WordPageSize.Unknown: return null;
                case WordPageSize.A6: return A6;
                case WordPageSize.B5: return B5;
                case WordPageSize.Letter: return Letter;
                case WordPageSize.Statement: return Statement;
                case WordPageSize.Legal: return Legal;
            }

            throw new ArgumentOutOfRangeException(nameof(pageSize));
        }

        public static PageSize A3 {
            get {
                return new PageSize() {
                    Width = (UInt32Value)16838U,
                    Height = (UInt32Value)23811U,
                    Code = (UInt16Value)8U
                };
            }
        }

        public static PageSize A4 {
            get {
                return new PageSize() {
                    Width = (UInt32Value)11906U,
                    Height = (UInt32Value)16838U,
                    Code = (UInt16Value)9U
                };
            }
        }

        public static PageSize A5 {
            get {
                return new PageSize() {
                    Width = (UInt32Value)8391U,
                    Height = (UInt32Value)11906U,
                    Code = (UInt16Value)11U
                };
            }
        }

        public static PageSize Executive =>
            new PageSize() {
                Width = (UInt32Value)10440U,
                Height = (UInt32Value)15120U,
                Code = (UInt16Value)7U
            };

        public static PageSize A6 => new PageSize() { Width = (UInt32Value)5953U, Height = (UInt32Value)8391U, Code = (UInt16Value)70U };
        public static PageSize B5 => new PageSize() { Width = (UInt32Value)10318U, Height = (UInt32Value)14570U, Code = (UInt16Value)13U };
        public static PageSize Statement => new PageSize() { Width = (UInt32Value)7920U, Height = (UInt32Value)12240U, Code = (UInt16Value)6U };
        public static PageSize Legal => new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)20160U, Code = (UInt16Value)5U };
        public static PageSize Letter => new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U, Code = (UInt16Value)1U };
    }
}