using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Authoring-time style theme that groups reusable shape, connector, and text styles.
    /// </summary>
    public sealed class VisioStyleTheme {
        /// <summary>Theme name.</summary>
        public string Name { get; set; } = "OfficeIMO Modern";

        /// <summary>Primary process/block shape style.</summary>
        public VisioShapeStyle Primary { get; set; } = Shape(Color.FromRgb(45, 145, 225), Color.FromRgb(26, 111, 188), 0.018, Color.White);

        /// <summary>Success/terminator shape style.</summary>
        public VisioShapeStyle Success { get; set; } = Shape(Color.FromRgb(122, 202, 94), Color.FromRgb(93, 166, 68), 0.018, Color.White);

        /// <summary>Decision/attention shape style.</summary>
        public VisioShapeStyle Decision { get; set; } = Shape(Color.FromRgb(120, 199, 89), Color.FromRgb(93, 166, 68), 0.018, Color.White);

        /// <summary>Continuation/marker shape style.</summary>
        public VisioShapeStyle Marker { get; set; } = Shape(Color.FromRgb(20, 119, 211), Color.FromRgb(16, 96, 170), 0.018, Color.White);

        /// <summary>Emphasis shape style.</summary>
        public VisioShapeStyle Emphasis { get; set; } = Shape(Color.FromRgb(88, 88, 88), Color.FromRgb(64, 64, 64), 0.018, Color.White);

        /// <summary>Container/region shape style.</summary>
        public VisioShapeStyle Container { get; set; } = Shape(Color.FromRgb(218, 242, 252), Color.FromRgb(146, 202, 224), 0.014, Color.FromRgb(34, 62, 80));

        /// <summary>Standard connector style.</summary>
        public VisioConnectorStyle Connector { get; set; } = CreateConnector(Color.FromRgb(45, 145, 225), 0.018, 1, Color.FromRgb(26, 111, 188));

        /// <summary>Data-flow connector style.</summary>
        public VisioConnectorStyle DataConnector { get; set; } = CreateConnector(Color.FromRgb(0, 146, 203), 0.018, 1, Color.FromRgb(0, 106, 160));

        /// <summary>Control-flow connector style.</summary>
        public VisioConnectorStyle ControlConnector { get; set; } = CreateConnector(Color.FromRgb(0, 146, 203), 0.018, 2, Color.FromRgb(0, 106, 160));

        /// <summary>Creates a detached copy of the theme.</summary>
        public VisioStyleTheme Clone() {
            return new VisioStyleTheme {
                Name = Name,
                Primary = Primary.Clone(),
                Success = Success.Clone(),
                Decision = Decision.Clone(),
                Marker = Marker.Clone(),
                Emphasis = Emphasis.Clone(),
                Container = Container.Clone(),
                Connector = Connector.Clone(),
                DataConnector = DataConnector.Clone(),
                ControlConnector = ControlConnector.Clone()
            };
        }

        /// <summary>Converts this authoring style theme into a flowchart builder theme.</summary>
        public OfficeIMO.Visio.Diagrams.VisioFlowchartTheme ToFlowchartTheme() {
            return new OfficeIMO.Visio.Diagrams.VisioFlowchartTheme {
                ProcessFill = Primary.FillColor,
                ProcessStroke = Primary.LineColor,
                DecisionFill = Decision.FillColor,
                DecisionStroke = Decision.LineColor,
                TerminatorFill = Success.FillColor,
                TerminatorStroke = Success.LineColor,
                MarkerFill = Marker.FillColor,
                MarkerStroke = Marker.LineColor,
                ConnectorColor = Connector.LineColor,
                LineWeight = Primary.LineWeight,
                ProcessTextStyle = Primary.TextStyle?.Clone(),
                DecisionTextStyle = Decision.TextStyle?.Clone(),
                TerminatorTextStyle = Success.TextStyle?.Clone(),
                MarkerTextStyle = Marker.TextStyle?.Clone(),
                ConnectorTextStyle = Connector.TextStyle?.Clone()
            };
        }

        /// <summary>Converts this authoring style theme into a block diagram builder theme.</summary>
        public OfficeIMO.Visio.Diagrams.VisioBlockDiagramTheme ToBlockDiagramTheme() {
            return new OfficeIMO.Visio.Diagrams.VisioBlockDiagramTheme {
                BlockFill = Primary.FillColor,
                BlockStroke = Primary.LineColor,
                EmphasisFill = Emphasis.FillColor,
                EmphasisStroke = Emphasis.LineColor,
                RegionFill = Container.FillColor,
                RegionStroke = Container.LineColor,
                DataFlowColor = DataConnector.LineColor,
                ControlFlowColor = ControlConnector.LineColor,
                LineWeight = Primary.LineWeight,
                BlockTextStyle = Primary.TextStyle?.Clone(),
                EmphasisTextStyle = Emphasis.TextStyle?.Clone(),
                RegionTextStyle = Container.TextStyle?.Clone(),
                ConnectorTextStyle = Connector.TextStyle?.Clone(),
                LegendTextStyle = DataConnector.TextStyle?.Clone()
            };
        }

        /// <summary>Polished blue/green theme for common business diagrams.</summary>
        public static VisioStyleTheme Modern() {
            return new VisioStyleTheme();
        }

        /// <summary>Office-like blue, green, and gold theme for familiar business diagrams.</summary>
        public static VisioStyleTheme Office() {
            return new VisioStyleTheme {
                Name = "OfficeIMO Office",
                Primary = Shape(Color.FromRgb(68, 114, 196), Color.FromRgb(47, 84, 150), 0.018, Color.White),
                Success = Shape(Color.FromRgb(112, 173, 71), Color.FromRgb(84, 130, 53), 0.018, Color.White),
                Decision = Shape(Color.FromRgb(255, 192, 0), Color.FromRgb(191, 144, 0), 0.018, Color.FromRgb(50, 42, 12)),
                Marker = Shape(Color.FromRgb(91, 155, 213), Color.FromRgb(47, 117, 181), 0.018, Color.White),
                Emphasis = Shape(Color.FromRgb(112, 48, 160), Color.FromRgb(80, 35, 115), 0.018, Color.White),
                Container = Shape(Color.FromRgb(221, 235, 247), Color.FromRgb(157, 195, 230), 0.014, Color.FromRgb(31, 78, 121)),
                Connector = CreateConnector(Color.FromRgb(68, 114, 196), 0.018, 1, Color.FromRgb(47, 84, 150)),
                DataConnector = CreateConnector(Color.FromRgb(91, 155, 213), 0.018, 1, Color.FromRgb(47, 117, 181)),
                ControlConnector = CreateConnector(Color.FromRgb(112, 48, 160), 0.018, 2, Color.FromRgb(80, 35, 115))
            };
        }

        /// <summary>Fluent-inspired theme using crisp Microsoft 365 accent colors.</summary>
        public static VisioStyleTheme Fluent() {
            return new VisioStyleTheme {
                Name = "OfficeIMO Fluent",
                Primary = Shape(Color.FromRgb(0, 120, 212), Color.FromRgb(0, 90, 158), 0.018, Color.White),
                Success = Shape(Color.FromRgb(16, 124, 16), Color.FromRgb(10, 92, 10), 0.018, Color.White),
                Decision = Shape(Color.FromRgb(255, 185, 0), Color.FromRgb(194, 124, 14), 0.018, Color.FromRgb(50, 42, 12)),
                Marker = Shape(Color.FromRgb(92, 45, 145), Color.FromRgb(68, 34, 107), 0.018, Color.White),
                Emphasis = Shape(Color.FromRgb(96, 94, 92), Color.FromRgb(72, 70, 68), 0.018, Color.White),
                Container = Shape(Color.FromRgb(239, 246, 252), Color.FromRgb(163, 205, 237), 0.014, Color.FromRgb(32, 80, 114)),
                Connector = CreateConnector(Color.FromRgb(0, 120, 212), 0.018, 1, Color.FromRgb(0, 90, 158)),
                DataConnector = CreateConnector(Color.FromRgb(0, 153, 188), 0.018, 1, Color.FromRgb(0, 103, 124)),
                ControlConnector = CreateConnector(Color.FromRgb(92, 45, 145), 0.018, 2, Color.FromRgb(68, 34, 107))
            };
        }

        /// <summary>Quiet technical theme for infrastructure and system diagrams.</summary>
        public static VisioStyleTheme Technical() {
            return new VisioStyleTheme {
                Name = "OfficeIMO Technical",
                Primary = Shape(Color.FromRgb(0, 146, 203), Color.FromRgb(0, 106, 160), 0.018, Color.White),
                Success = Shape(Color.FromRgb(72, 178, 112), Color.FromRgb(46, 125, 79), 0.018, Color.White),
                Decision = Shape(Color.FromRgb(245, 183, 77), Color.FromRgb(181, 119, 24), 0.018, Color.FromRgb(54, 41, 12)),
                Marker = Shape(Color.FromRgb(0, 120, 212), Color.FromRgb(0, 90, 158), 0.018, Color.White),
                Emphasis = Shape(Color.FromRgb(88, 88, 88), Color.FromRgb(64, 64, 64), 0.018, Color.White),
                Container = Shape(Color.FromRgb(218, 242, 252), Color.FromRgb(146, 202, 224), 0.014, Color.FromRgb(34, 62, 80)),
                Connector = CreateConnector(Color.FromRgb(0, 120, 212), 0.018, 1, Color.FromRgb(0, 90, 158)),
                DataConnector = CreateConnector(Color.FromRgb(0, 146, 203), 0.018, 1, Color.FromRgb(0, 106, 160)),
                ControlConnector = CreateConnector(Color.FromRgb(0, 146, 203), 0.018, 2, Color.FromRgb(0, 106, 160))
            };
        }

        /// <summary>Low-ink theme for clean documentation and printed pages.</summary>
        public static VisioStyleTheme Minimal() {
            return new VisioStyleTheme {
                Name = "OfficeIMO Minimal",
                Primary = Shape(Color.White, Color.FromRgb(80, 91, 105), 0.014, Color.FromRgb(31, 41, 55)),
                Success = Shape(Color.FromRgb(237, 247, 239), Color.FromRgb(67, 137, 84), 0.014, Color.FromRgb(31, 88, 45)),
                Decision = Shape(Color.FromRgb(255, 248, 225), Color.FromRgb(166, 124, 0), 0.014, Color.FromRgb(92, 66, 0)),
                Marker = Shape(Color.FromRgb(239, 246, 255), Color.FromRgb(37, 99, 235), 0.014, Color.FromRgb(30, 64, 175)),
                Emphasis = Shape(Color.FromRgb(245, 247, 250), Color.FromRgb(80, 91, 105), 0.014, Color.FromRgb(31, 41, 55)),
                Container = Shape(Color.FromRgb(248, 250, 252), Color.FromRgb(203, 213, 225), 0.01, Color.FromRgb(51, 65, 85)),
                Connector = CreateConnector(Color.FromRgb(80, 91, 105), 0.014, 1, Color.FromRgb(51, 65, 85)),
                DataConnector = CreateConnector(Color.FromRgb(37, 99, 235), 0.014, 1, Color.FromRgb(30, 64, 175)),
                ControlConnector = CreateConnector(Color.FromRgb(80, 91, 105), 0.014, 2, Color.FromRgb(51, 65, 85))
            };
        }

        /// <summary>Dark presentation theme with high-contrast labels for dashboards and executive diagrams.</summary>
        public static VisioStyleTheme Dark() {
            return new VisioStyleTheme {
                Name = "OfficeIMO Dark",
                Primary = Shape(Color.FromRgb(37, 99, 235), Color.FromRgb(96, 165, 250), 0.018, Color.White),
                Success = Shape(Color.FromRgb(5, 150, 105), Color.FromRgb(52, 211, 153), 0.018, Color.White),
                Decision = Shape(Color.FromRgb(217, 119, 6), Color.FromRgb(251, 191, 36), 0.018, Color.White),
                Marker = Shape(Color.FromRgb(14, 116, 144), Color.FromRgb(103, 232, 249), 0.018, Color.White),
                Emphasis = Shape(Color.FromRgb(51, 65, 85), Color.FromRgb(148, 163, 184), 0.018, Color.White),
                Container = Shape(Color.FromRgb(15, 23, 42), Color.FromRgb(71, 85, 105), 0.014, Color.FromRgb(226, 232, 240)),
                Connector = CreateConnector(Color.FromRgb(96, 165, 250), 0.018, 1, Color.FromRgb(191, 219, 254)),
                DataConnector = CreateConnector(Color.FromRgb(34, 211, 238), 0.018, 1, Color.FromRgb(207, 250, 254)),
                ControlConnector = CreateConnector(Color.FromRgb(196, 181, 253), 0.018, 2, Color.FromRgb(237, 233, 254))
            };
        }

        /// <summary>Black-and-white theme for maximum print compatibility.</summary>
        public static VisioStyleTheme Print() {
            return new VisioStyleTheme {
                Name = "OfficeIMO Print",
                Primary = Shape(Color.White, Color.Black, 0.012, Color.Black),
                Success = Shape(Color.White, Color.Black, 0.012, Color.Black),
                Decision = Shape(Color.White, Color.Black, 0.012, Color.Black),
                Marker = Shape(Color.White, Color.Black, 0.012, Color.Black),
                Emphasis = Shape(Color.FromRgb(235, 235, 235), Color.Black, 0.012, Color.Black),
                Container = Shape(Color.White, Color.FromRgb(128, 128, 128), 0.008, Color.Black),
                Connector = CreateConnector(Color.Black, 0.012, 1, Color.Black),
                DataConnector = CreateConnector(Color.Black, 0.012, 1, Color.Black),
                ControlConnector = CreateConnector(Color.Black, 0.012, 2, Color.Black)
            };
        }

        private static VisioShapeStyle Shape(Color fill, Color line, double weight, Color text, int linePattern = 1, int fillPattern = 1) {
            return new VisioShapeStyle(fill, line, weight, linePattern, fillPattern) {
                TextStyle = Text(text)
            };
        }

        private static VisioConnectorStyle CreateConnector(Color line, double weight, int pattern, Color text) {
            return new VisioConnectorStyle(line, weight, pattern, EndArrow.Triangle) {
                Kind = ConnectorKind.RightAngle,
                TextStyle = ConnectorText(text)
            };
        }

        private static VisioTextStyle Text(Color color) {
            return new VisioTextStyle {
                FontFamily = "Aptos",
                Color = color,
                Size = 10,
                Bold = true,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
        }

        private static VisioTextStyle ConnectorText(Color color) {
            return new VisioTextStyle {
                FontFamily = "Aptos",
                Color = color,
                Size = 9,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center
            };
        }
    }
}
