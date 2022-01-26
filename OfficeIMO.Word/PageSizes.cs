using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static class PageSizes {
        public static PageSize A3 {
            get {
                return new PageSize() {
                    Width = (UInt32Value) 16838U,
                    Height = (UInt32Value) 23811U,
                    Code = (UInt16Value) 8U
                };
            }
        }

        public static PageSize A4 {
            get {
                return new PageSize() {
                    Width = (UInt32Value) 11906U,
                    Height = (UInt32Value) 16838U,
                    Code = (UInt16Value) 9U
                };
            }
        }

        public static PageSize A5 {
            get {
                return new PageSize() {
                    Width = (UInt32Value) 8391U,
                    Height = (UInt32Value) 11906U,
                    Code = (UInt16Value) 11U
                };
            }
        }

        public static PageSize Executive {
            get {
                return new PageSize() {
                    Width = (UInt32Value) 10440U,
                    Height = (UInt32Value) 15120U,
                    Code = (UInt16Value) 7U
                };
            }
        }
    }
}