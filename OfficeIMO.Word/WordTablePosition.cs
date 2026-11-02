using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    ///     Positions a floating table using physical lengths and anchor alignments.
    /// </summary>
    /// <remarks>
    /// Lengths round to whole twips, with midpoint ties away from zero. Text gaps retain
    /// signed 16-bit storage; X and Y offsets retain signed 32-bit storage. A rounded
    /// value outside its native range throws before XML changes. Null removes the
    /// attribute without creating positioning properties; zero remains explicit.
    /// </remarks>
    public class WordTablePosition {
        private readonly WordTable _table;

        private TableProperties? TableProperties => _table._tableProperties;

        private TableProperties EnsureTableProperties() {
            _table.CheckTableProperties();
            return _table._tableProperties ?? throw new InvalidOperationException("Table properties are not available.");
        }

        /// <summary>
        ///     Constructor for Table Positioning
        /// </summary>
        /// <param name="table"></param>
        internal WordTablePosition(WordTable table) {
            _table = table ?? throw new ArgumentNullException(nameof(table));
        }

        /// <summary>
        ///     Get or set Distance From Left of Table to Text
        /// </summary>
        public OfficeLength? LeftFromText {
            get {
                var native = TableProperties?.TablePositionProperties?.LeftFromText;
                return native != null ? OfficeLength.Emus((long)native.Value * 635) : (OfficeLength?)null;
            }
            set {
                var native = checked((short?)(value?.ToTwips()));
                if (!native.HasValue) {
                    var position = TableProperties?.TablePositionProperties;
                    if (position != null) position.LeftFromText = null;
                    return;
                }

                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();
                tableProperties.TablePositionProperties.LeftFromText = native.Value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Right of Table to Text
        /// </summary>
        public OfficeLength? RightFromText {
            get {
                var native = TableProperties?.TablePositionProperties?.RightFromText;
                return native != null ? OfficeLength.Emus((long)native.Value * 635) : (OfficeLength?)null;
            }
            set {
                var native = checked((short?)(value?.ToTwips()));
                if (!native.HasValue) {
                    var position = TableProperties?.TablePositionProperties;
                    if (position != null) position.RightFromText = null;
                    return;
                }

                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();
                tableProperties.TablePositionProperties.RightFromText = native.Value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Bottom of Table to Text
        /// </summary>
        public OfficeLength? BottomFromText {
            get {
                var native = TableProperties?.TablePositionProperties?.BottomFromText;
                return native != null ? OfficeLength.Emus((long)native.Value * 635) : (OfficeLength?)null;
            }
            set {
                var native = checked((short?)(value?.ToTwips()));
                if (!native.HasValue) {
                    var position = TableProperties?.TablePositionProperties;
                    if (position != null) position.BottomFromText = null;
                    return;
                }

                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();
                tableProperties.TablePositionProperties.BottomFromText = native.Value;
            }
        }

        /// <summary>
        ///     Get or set Distance From Top of Table to Text
        /// </summary>
        public OfficeLength? TopFromText {
            get {
                var native = TableProperties?.TablePositionProperties?.TopFromText;
                return native != null ? OfficeLength.Emus((long)native.Value * 635) : (OfficeLength?)null;
            }
            set {
                var native = checked((short?)(value?.ToTwips()));
                if (!native.HasValue) {
                    var position = TableProperties?.TablePositionProperties;
                    if (position != null) position.TopFromText = null;
                    return;
                }

                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();
                tableProperties.TablePositionProperties.TopFromText = native.Value;
            }
        }

        /// <summary>
        ///     Get or set Table Vertical Anchor
        /// </summary>
        public WordTableVerticalAnchor? VerticalAnchor {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.VerticalAnchor != null)
                    return tableProperties.TablePositionProperties.VerticalAnchor.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.VerticalAnchor = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.VerticalAnchor = null;
            }
        }

        /// <summary>
        ///     Get or set Table Horizontal Anchor
        /// </summary>
        public WordTableHorizontalAnchor? HorizontalAnchor {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.HorizontalAnchor != null)
                    return tableProperties.TablePositionProperties.HorizontalAnchor.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.HorizontalAnchor = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.HorizontalAnchor = null;
            }
        }

        /// <summary>
        ///     Gets or sets the signed vertical distance from the anchor.
        /// </summary>
        public OfficeLength? TablePositionY {
            get {
                var native = TableProperties?.TablePositionProperties?.TablePositionY;
                return native != null ? OfficeLength.Emus((long)native.Value * 635) : (OfficeLength?)null;
            }
            set {
                var native = checked((int?)(value?.ToTwips()));
                if (!native.HasValue) {
                    var position = TableProperties?.TablePositionProperties;
                    if (position != null) position.TablePositionY = null;
                    return;
                }

                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();
                tableProperties.TablePositionProperties.TablePositionY = native.Value;
            }
        }

        /// <summary>
        ///     Get or set Absolute Horizontal Distance From Anchor
        /// </summary>
        public OfficeLength? TablePositionX {
            get {
                var native = TableProperties?.TablePositionProperties?.TablePositionX;
                return native != null ? OfficeLength.Emus((long)native.Value * 635) : (OfficeLength?)null;
            }
            set {
                var native = checked((int?)(value?.ToTwips()));
                if (!native.HasValue) {
                    var position = TableProperties?.TablePositionProperties;
                    if (position != null) position.TablePositionX = null;
                    return;
                }

                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();
                tableProperties.TablePositionProperties.TablePositionX = native.Value;
            }
        }

        /// <summary>
        ///     Get or set Relative Vertical Alignment from Anchor
        /// </summary>
        public WordTableVerticalPositionAlignment? TablePositionYAlignment {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionYAlignment != null)
                    return tableProperties.TablePositionProperties.TablePositionYAlignment.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionYAlignment = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.TablePositionYAlignment = null;
            }
        }

        /// <summary>
        ///     Get or set Relative Horizontal Alignment From Anchor
        /// </summary>
        public WordTableHorizontalAlignment? TablePositionXAlignment {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TablePositionProperties?.TablePositionXAlignment != null)
                    return tableProperties.TablePositionProperties.TablePositionXAlignment.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TablePositionProperties == null)
                    tableProperties.TablePositionProperties = new TablePositionProperties();

                if (value != null)
                    tableProperties.TablePositionProperties.TablePositionXAlignment = value.Value.ToOpenXml();
                else
                    tableProperties.TablePositionProperties.TablePositionXAlignment = null;
            }
        }

        /// <summary>
        ///     Gets or sets Table Overlap
        /// </summary>
        public WordTableOverlap? TableOverlap {
            get {
                var tableProperties = TableProperties;
                if (tableProperties?.TableOverlap?.Val != null)
                    return tableProperties.TableOverlap.Val.Value.ToOfficeEnum();

                return null;
            }
            set {
                var tableProperties = EnsureTableProperties();
                if (tableProperties.TableOverlap == null) tableProperties.TableOverlap = new TableOverlap();
                if (value != null)
                    tableProperties.TableOverlap.Val = value.Value.ToOpenXml();
                else
                    tableProperties.TableOverlap.Remove();
            }
        }
    }
}
