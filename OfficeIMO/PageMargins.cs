using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public static class PageMargins {
        public static PageMargin Normal {
            get {
                return new PageMargin() {
                    Top = 1440,
                    Right = (UInt32Value) 1440U,
                    Bottom = 1440,
                    Left = (UInt32Value) 1440U,
                    Header = (UInt32Value) 720U,
                    Footer = (UInt32Value) 720U,
                    Gutter = (UInt32Value) 0U
                };
            }
        }

        public static PageMargin Mirrored {
            get {
                return new PageMargin() {
                    Top = 1440,
                    Right = (UInt32Value) 1440U,
                    Bottom = 1440,
                    Left = (UInt32Value) 1800U,
                    Header = (UInt32Value) 720U,
                    Footer = (UInt32Value) 720U,
                    Gutter = (UInt32Value) 0U
                };
            }
        }

        public static PageMargin Moderate {
            get {
                return new PageMargin() {
                    Top = 1440, Right = (UInt32Value) 1080U,
                    Bottom = 1440, Left = (UInt32Value) 1080U,
                    Header = (UInt32Value) 720U,
                    Footer = (UInt32Value) 720U,
                    Gutter = (UInt32Value) 0U
                };
            }
        }

        public static PageMargin Narrow {
            get {
                return new PageMargin() {
                    Top = 720, Right = (UInt32Value) 720U,
                    Bottom = 720, Left = (UInt32Value) 720U,
                    Header = (UInt32Value) 720U,
                    Footer = (UInt32Value) 720U,
                    Gutter = (UInt32Value) 0U
                };
            }
        }

        public static PageMargin Wide {
            get {
                return new PageMargin() {
                    Top = 1440, Right = (UInt32Value) 2880U,
                    Bottom = 1440, Left = (UInt32Value) 2880U,
                    Header = (UInt32Value) 720U,
                    Footer = (UInt32Value) 720U,
                    Gutter = (UInt32Value) 0U
                };
            }
        }
    }
}