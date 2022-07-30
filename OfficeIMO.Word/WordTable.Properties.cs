using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        public enum AutoFitType {
            ToContent,
            ToWindow,
            Fixed
        }

        public AutoFitType? AutoFit {
            get {
                return AutoFitType.ToContent;
            }
        }

        /// <summary>
        /// Gets or sets a Title/Caption to a Table
        /// </summary>
        public string Title {
            get {
                if (_tableProperties != null && _tableProperties.TableCaption != null) {
                    return _tableProperties.TableCaption.Val;
                }

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableCaption == null) {
                    _tableProperties.TableCaption = new TableCaption();
                }
                if (value != null) {
                    _tableProperties.TableCaption.Val = value;
                } else {
                    _tableProperties.TableCaption.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets Description for a Table
        /// </summary>
        public string Description {
            get {
                if (_tableProperties != null && _tableProperties.TableDescription != null) {
                    return _tableProperties.TableDescription.Val;
                }

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableDescription == null) {
                    _tableProperties.TableDescription = new TableDescription();
                }
                if (value != null) {
                    _tableProperties.TableDescription.Val = value;
                } else {
                    _tableProperties.TableDescription.Remove();
                }
            }
        }

        /// <summary>
        /// Allow table to overlap or not
        /// </summary>
        public bool AllowOverlap {
            get {
                if (this.Position.TableOverlap == TableOverlapValues.Overlap) {
                    return true;
                }
                return false;
            }
            set => this.Position.TableOverlap = value ? TableOverlapValues.Overlap : TableOverlapValues.Never;
        }

        /// <summary>
        /// Allow text to wrap around table.
        /// </summary>
        public bool AllowTextWrap {
            get {
                if (this.Position.VerticalAnchor == VerticalAnchorValues.Text) {
                    return true;
                }

                return false;
            }
            set {
                if (value == true) {
                    this.Position.VerticalAnchor = VerticalAnchorValues.Text;
                } else {
                    this.Position.VerticalAnchor = null;
                }
            }
        }

    }
}
