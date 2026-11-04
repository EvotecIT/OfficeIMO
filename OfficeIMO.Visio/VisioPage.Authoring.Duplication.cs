using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio {
    public partial class VisioPage {
        private const string DefaultDuplicateIdSuffix = "-copy";

        /// <summary>
        /// Duplicates one shape by id, copies its grouped children, and keeps copied ids easy to reference.
        /// </summary>
        /// <param name="shapeId">Identifier of the shape to duplicate.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioPage DuplicateShape(string shapeId, Action<VisioShapeSelection>? configure = null) {
            return DuplicateShape(shapeId, CreateDefaultDuplicationOptions(), configure);
        }

        /// <summary>
        /// Duplicates one shape by id using explicit duplication options.
        /// </summary>
        /// <param name="shapeId">Identifier of the shape to duplicate.</param>
        /// <param name="options">Duplication options.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioPage DuplicateShape(string shapeId, VisioShapeDuplicationOptions? options, Action<VisioShapeSelection>? configure = null) {
            return DuplicateShapes(new[] { shapeId }, options, configure);
        }

        /// <summary>
        /// Duplicates several shapes by id, preserving internal connectors between duplicated shapes.
        /// </summary>
        /// <param name="shapeIds">Shape identifiers to duplicate.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioPage DuplicateShapes(IEnumerable<string> shapeIds, Action<VisioShapeSelection>? configure = null) {
            return DuplicateShapes(shapeIds, CreateDefaultDuplicationOptions(), configure);
        }

        /// <summary>
        /// Duplicates several shapes by id using explicit duplication options.
        /// </summary>
        /// <param name="shapeIds">Shape identifiers to duplicate.</param>
        /// <param name="options">Duplication options.</param>
        /// <param name="configure">Optional configuration for the duplicated selection.</param>
        public VisioPage DuplicateShapes(IEnumerable<string> shapeIds, VisioShapeDuplicationOptions? options, Action<VisioShapeSelection>? configure = null) {
            ArgumentNullException.ThrowIfNull(shapeIds);

            List<VisioShape> shapes = new();
            foreach (string shapeId in shapeIds) {
                shapes.Add(GetShape(shapeId));
            }

            VisioShapeSelection duplicates = this.DuplicateShapes(shapes, options);
            configure?.Invoke(duplicates);
            ValidateShapeIdentifiers();
            return this;
        }

        private static VisioShapeDuplicationOptions CreateDefaultDuplicationOptions() {
            return new VisioShapeDuplicationOptions {
                IdSuffix = DefaultDuplicateIdSuffix,
                ConnectorIdSuffix = DefaultDuplicateIdSuffix
            };
        }
    }
}
