namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one OfficeArt rule connecting two shapes through a connector shape.</summary>
    public sealed class LegacyPptConnectorRule {
        internal LegacyPptConnectorRule(uint ruleId, uint startShapeId, uint endShapeId,
            uint connectorShapeId, uint startConnectionSiteIndex, uint endConnectionSiteIndex) {
            RuleId = ruleId;
            StartShapeId = startShapeId;
            EndShapeId = endShapeId;
            ConnectorShapeId = connectorShapeId;
            StartConnectionSiteIndex = startConnectionSiteIndex;
            EndConnectionSiteIndex = endConnectionSiteIndex;
        }

        /// <summary>Gets the OfficeArt solver-rule identifier.</summary>
        public uint RuleId { get; }

        /// <summary>Gets the OfficeArt shape identifier at the connector start.</summary>
        public uint StartShapeId { get; }

        /// <summary>Gets the OfficeArt shape identifier at the connector end.</summary>
        public uint EndShapeId { get; }

        /// <summary>Gets the OfficeArt identifier of the connector shape.</summary>
        public uint ConnectorShapeId { get; }

        /// <summary>Gets the connection-site index on the start shape.</summary>
        public uint StartConnectionSiteIndex { get; }

        /// <summary>Gets the connection-site index on the end shape.</summary>
        public uint EndConnectionSiteIndex { get; }
    }
}
