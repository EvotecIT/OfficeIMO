using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        private const string DefaultFluentDuplicateIdSuffix = "-copy";

        /// <summary>
        /// Duplicates one shape by id, copies its grouped children, and keeps copied ids easy to reference.
        /// </summary>
        /// <param name="shapeId">Identifier of the shape to duplicate.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioFluentPage DuplicateShape(string shapeId, Action<VisioShapeSelection>? configure = null) {
            return DuplicateShape(shapeId, CreateDefaultFluentDuplicationOptions(), configure);
        }

        /// <summary>
        /// Duplicates one shape by id using explicit duplication options.
        /// </summary>
        /// <param name="shapeId">Identifier of the shape to duplicate.</param>
        /// <param name="options">Duplication options.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioFluentPage DuplicateShape(string shapeId, VisioShapeDuplicationOptions? options, Action<VisioShapeSelection>? configure = null) {
            return DuplicateShapes(new[] { shapeId }, options, configure);
        }

        /// <summary>
        /// Duplicates several shapes by id, preserving internal connectors between duplicated shapes.
        /// </summary>
        /// <param name="shapeIds">Shape identifiers to duplicate.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioFluentPage DuplicateShapes(IEnumerable<string> shapeIds, Action<VisioShapeSelection>? configure = null) {
            return DuplicateShapes(shapeIds, CreateDefaultFluentDuplicationOptions(), configure);
        }

        /// <summary>
        /// Duplicates several shapes by id using explicit duplication options.
        /// </summary>
        /// <param name="shapeIds">Shape identifiers to duplicate.</param>
        /// <param name="options">Duplication options.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioFluentPage DuplicateShapes(IEnumerable<string> shapeIds, VisioShapeDuplicationOptions? options, Action<VisioShapeSelection>? configure = null) {
            if (shapeIds == null) {
                throw new ArgumentNullException(nameof(shapeIds));
            }

            List<VisioShape> shapes = new();
            foreach (string shapeId in shapeIds) {
                shapes.Add(ResolveShape(shapeId));
            }

            VisioShapeSelection duplicates = Page.DuplicateShapes(shapes, options);
            configure?.Invoke(duplicates);
            RebuildShapeIndex();
            return this;
        }

        private static VisioShapeDuplicationOptions CreateDefaultFluentDuplicationOptions() {
            return new VisioShapeDuplicationOptions {
                IdSuffix = DefaultFluentDuplicateIdSuffix,
                ConnectorIdSuffix = DefaultFluentDuplicateIdSuffix
            };
        }
    }
}
