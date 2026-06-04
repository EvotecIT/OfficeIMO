using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Match strategy used by a stencil migration rule.
    /// </summary>
    public enum VisioStencilMigrationMatchKind {
        /// <summary>Match by the current shape master universal name.</summary>
        MasterNameU,

        /// <summary>Match by the current shape universal name.</summary>
        ShapeNameU,

        /// <summary>Match by the OfficeIMO stencil id stored on the shape.</summary>
        StencilId,

        /// <summary>Match by a caller-provided predicate.</summary>
        Predicate
    }

    /// <summary>
    /// Declarative map for replacing existing shapes with first-party or package-backed stencil definitions.
    /// </summary>
    public sealed class VisioStencilMigrationMap {
        private readonly IReadOnlyList<VisioStencilMigrationRule> _rules;

        internal VisioStencilMigrationMap(IEnumerable<VisioStencilMigrationRule> rules) {
            _rules = rules.ToList().AsReadOnly();
        }

        /// <summary>
        /// Gets migration rules in first-match-wins order.
        /// </summary>
        public IReadOnlyList<VisioStencilMigrationRule> Rules => _rules;

        /// <summary>
        /// Creates a migration map with a fluent builder.
        /// </summary>
        /// <param name="configure">Builder configuration.</param>
        public static VisioStencilMigrationMap Create(Action<VisioStencilMigrationMapBuilder> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            VisioStencilMigrationMapBuilder builder = new();
            configure(builder);
            return builder.Build();
        }

        internal VisioStencilMigrationRule? FindRule(VisioShape shape) {
            foreach (VisioStencilMigrationRule rule in _rules) {
                if (rule.IsMatch(shape)) {
                    return rule;
                }
            }

            return null;
        }
    }

    /// <summary>
    /// Builds a <see cref="VisioStencilMigrationMap"/>.
    /// </summary>
    public sealed class VisioStencilMigrationMapBuilder {
        private readonly List<VisioStencilMigrationRule> _rules = new();

        /// <summary>
        /// Maps shapes with the current master universal name to a replacement stencil.
        /// </summary>
        public VisioStencilMigrationMapBuilder MapMaster(string currentMasterNameU, VisioStencilShape replacement, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            _rules.Add(VisioStencilMigrationRule.ForMaster(currentMasterNameU, replacement, resizeToStencil, comparison));
            return this;
        }

        /// <summary>
        /// Maps shapes with the current shape universal name to a replacement stencil.
        /// </summary>
        public VisioStencilMigrationMapBuilder MapNameU(string currentNameU, VisioStencilShape replacement, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            _rules.Add(VisioStencilMigrationRule.ForNameU(currentNameU, replacement, resizeToStencil, comparison));
            return this;
        }

        /// <summary>
        /// Maps shapes carrying the current OfficeIMO stencil id to a replacement stencil.
        /// </summary>
        public VisioStencilMigrationMapBuilder MapStencilId(string currentStencilId, VisioStencilShape replacement, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            _rules.Add(VisioStencilMigrationRule.ForStencilId(currentStencilId, replacement, resizeToStencil, comparison));
            return this;
        }

        /// <summary>
        /// Maps shapes accepted by a typed predicate to a replacement stencil.
        /// </summary>
        public VisioStencilMigrationMapBuilder Map(Func<VisioShape, bool> predicate, VisioStencilShape replacement, bool resizeToStencil = false) {
            _rules.Add(VisioStencilMigrationRule.ForPredicate(predicate, replacement, resizeToStencil));
            return this;
        }

        /// <summary>
        /// Builds the migration map.
        /// </summary>
        public VisioStencilMigrationMap Build() {
            return new VisioStencilMigrationMap(_rules);
        }
    }

    /// <summary>
    /// One stencil migration rule.
    /// </summary>
    public sealed class VisioStencilMigrationRule {
        private readonly Func<VisioShape, bool>? _predicate;

        private VisioStencilMigrationRule(
            VisioStencilMigrationMatchKind matchKind,
            string? matchValue,
            Func<VisioShape, bool>? predicate,
            VisioStencilShape replacement,
            bool resizeToStencil,
            StringComparison comparison) {
            MatchKind = matchKind;
            MatchValue = matchValue;
            _predicate = predicate;
            Replacement = replacement ?? throw new ArgumentNullException(nameof(replacement));
            ResizeToStencil = resizeToStencil;
            Comparison = comparison;
        }

        /// <summary>Gets the match strategy.</summary>
        public VisioStencilMigrationMatchKind MatchKind { get; }

        /// <summary>Gets the match value for non-predicate rules.</summary>
        public string? MatchValue { get; }

        /// <summary>Gets the replacement stencil definition.</summary>
        public VisioStencilShape Replacement { get; }

        /// <summary>Gets whether matching shapes should be resized to the replacement stencil default size.</summary>
        public bool ResizeToStencil { get; }

        /// <summary>Gets the string comparison used by non-predicate rules.</summary>
        public StringComparison Comparison { get; }

        /// <summary>
        /// Creates a rule matching by current master universal name.
        /// </summary>
        public static VisioStencilMigrationRule ForMaster(string currentMasterNameU, VisioStencilShape replacement, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioStencilMigrationRule(VisioStencilMigrationMatchKind.MasterNameU, RequireValue(currentMasterNameU, nameof(currentMasterNameU)), null, replacement, resizeToStencil, comparison);
        }

        /// <summary>
        /// Creates a rule matching by current shape universal name.
        /// </summary>
        public static VisioStencilMigrationRule ForNameU(string currentNameU, VisioStencilShape replacement, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioStencilMigrationRule(VisioStencilMigrationMatchKind.ShapeNameU, RequireValue(currentNameU, nameof(currentNameU)), null, replacement, resizeToStencil, comparison);
        }

        /// <summary>
        /// Creates a rule matching by current OfficeIMO stencil id metadata.
        /// </summary>
        public static VisioStencilMigrationRule ForStencilId(string currentStencilId, VisioStencilShape replacement, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioStencilMigrationRule(VisioStencilMigrationMatchKind.StencilId, RequireValue(currentStencilId, nameof(currentStencilId)), null, replacement, resizeToStencil, comparison);
        }

        /// <summary>
        /// Creates a rule matching by caller-provided predicate.
        /// </summary>
        public static VisioStencilMigrationRule ForPredicate(Func<VisioShape, bool> predicate, VisioStencilShape replacement, bool resizeToStencil = false) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            return new VisioStencilMigrationRule(VisioStencilMigrationMatchKind.Predicate, null, predicate, replacement, resizeToStencil, StringComparison.OrdinalIgnoreCase);
        }

        internal bool IsMatch(VisioShape shape) {
            if (shape == null) {
                return false;
            }

            switch (MatchKind) {
                case VisioStencilMigrationMatchKind.MasterNameU:
                    return string.Equals(shape.MasterNameU, MatchValue, Comparison);
                case VisioStencilMigrationMatchKind.ShapeNameU:
                    return string.Equals(shape.NameU, MatchValue, Comparison);
                case VisioStencilMigrationMatchKind.StencilId:
                    return string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.StencilId), MatchValue, Comparison);
                case VisioStencilMigrationMatchKind.Predicate:
                    return _predicate != null && _predicate(shape);
                default:
                    return false;
            }
        }

        private static string RequireValue(string value, string parameterName) {
            if (string.IsNullOrWhiteSpace(value)) {
                throw new ArgumentException("Migration match value cannot be null or whitespace.", parameterName);
            }

            return value;
        }
    }
}
