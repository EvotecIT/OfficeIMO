using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Catalog-query helpers for building stencil migration maps without manually resolving every replacement shape.
    /// </summary>
    public static class VisioStencilMigrationMapBuilderExtensions {
        /// <summary>
        /// Maps shapes with the current master universal name to the best matching replacement stencil in a catalog.
        /// </summary>
        public static VisioStencilMigrationMapBuilder MapMaster(this VisioStencilMigrationMapBuilder builder, string currentMasterNameU, VisioStencilCatalog catalog, string replacementQuery, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return MapMaster(builder, currentMasterNameU, catalog, new[] { replacementQuery }, resizeToStencil, comparison);
        }

        /// <summary>
        /// Maps shapes with the current master universal name to the best matching replacement stencil in a catalog.
        /// Queries are tried in order and exact id/name/master/keyword/alias/tag matches win before search matches.
        /// </summary>
        public static VisioStencilMigrationMapBuilder MapMaster(this VisioStencilMigrationMapBuilder builder, string currentMasterNameU, VisioStencilCatalog catalog, IEnumerable<string> replacementQueries, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            RequireBuilder(builder);
            return builder.MapMaster(currentMasterNameU, Resolve(catalog, replacementQueries, nameof(replacementQueries)), resizeToStencil, comparison);
        }

        /// <summary>
        /// Maps shapes with the current shape universal name to the best matching replacement stencil in a catalog.
        /// </summary>
        public static VisioStencilMigrationMapBuilder MapNameU(this VisioStencilMigrationMapBuilder builder, string currentNameU, VisioStencilCatalog catalog, string replacementQuery, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return MapNameU(builder, currentNameU, catalog, new[] { replacementQuery }, resizeToStencil, comparison);
        }

        /// <summary>
        /// Maps shapes with the current shape universal name to the best matching replacement stencil in a catalog.
        /// Queries are tried in order and exact id/name/master/keyword/alias/tag matches win before search matches.
        /// </summary>
        public static VisioStencilMigrationMapBuilder MapNameU(this VisioStencilMigrationMapBuilder builder, string currentNameU, VisioStencilCatalog catalog, IEnumerable<string> replacementQueries, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            RequireBuilder(builder);
            return builder.MapNameU(currentNameU, Resolve(catalog, replacementQueries, nameof(replacementQueries)), resizeToStencil, comparison);
        }

        /// <summary>
        /// Maps shapes carrying the current OfficeIMO stencil id to the best matching replacement stencil in a catalog.
        /// </summary>
        public static VisioStencilMigrationMapBuilder MapStencilId(this VisioStencilMigrationMapBuilder builder, string currentStencilId, VisioStencilCatalog catalog, string replacementQuery, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return MapStencilId(builder, currentStencilId, catalog, new[] { replacementQuery }, resizeToStencil, comparison);
        }

        /// <summary>
        /// Maps shapes carrying the current OfficeIMO stencil id to the best matching replacement stencil in a catalog.
        /// Queries are tried in order and exact id/name/master/keyword/alias/tag matches win before search matches.
        /// </summary>
        public static VisioStencilMigrationMapBuilder MapStencilId(this VisioStencilMigrationMapBuilder builder, string currentStencilId, VisioStencilCatalog catalog, IEnumerable<string> replacementQueries, bool resizeToStencil = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            RequireBuilder(builder);
            return builder.MapStencilId(currentStencilId, Resolve(catalog, replacementQueries, nameof(replacementQueries)), resizeToStencil, comparison);
        }

        /// <summary>
        /// Maps shapes accepted by a typed predicate to the best matching replacement stencil in a catalog.
        /// </summary>
        public static VisioStencilMigrationMapBuilder Map(this VisioStencilMigrationMapBuilder builder, Func<VisioShape, bool> predicate, VisioStencilCatalog catalog, string replacementQuery, bool resizeToStencil = false) {
            return Map(builder, predicate, catalog, new[] { replacementQuery }, resizeToStencil);
        }

        /// <summary>
        /// Maps shapes accepted by a typed predicate to the best matching replacement stencil in a catalog.
        /// Queries are tried in order and exact id/name/master/keyword/alias/tag matches win before search matches.
        /// </summary>
        public static VisioStencilMigrationMapBuilder Map(this VisioStencilMigrationMapBuilder builder, Func<VisioShape, bool> predicate, VisioStencilCatalog catalog, IEnumerable<string> replacementQueries, bool resizeToStencil = false) {
            RequireBuilder(builder);
            return builder.Map(predicate, Resolve(catalog, replacementQueries, nameof(replacementQueries)), resizeToStencil);
        }

        private static void RequireBuilder(VisioStencilMigrationMapBuilder builder) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }
        }

        internal static VisioStencilShape Resolve(VisioStencilCatalog catalog, IEnumerable<string> replacementQueries, string parameterName) {
            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            if (replacementQueries == null) {
                throw new ArgumentNullException(parameterName);
            }

            string[] queries = replacementQueries
                .Where(query => !string.IsNullOrWhiteSpace(query))
                .ToArray();
            if (queries.Length == 0) {
                throw new ArgumentException("At least one stencil replacement query is required.", parameterName);
            }

            return catalog.FindBest(queries);
        }
    }
}
