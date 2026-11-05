using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {
        /// <summary>
        /// Adds existing shapes to a Visio-native container.
        /// </summary>
        public VisioPage AddToContainer(string containerId, IEnumerable<string> memberIds, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            VisioShape container = GetShape(containerId);
            this.AddToContainer(container, ResolveShapes(memberIds, nameof(memberIds)), resizeToFit, resizeOptions);
            return this;
        }

        /// <summary>
        /// Adds existing shapes to a Visio-native container using inline resize option configuration.
        /// </summary>
        public VisioPage AddToContainer(string containerId, IEnumerable<string> memberIds, Action<VisioContainerOptions> configureResizeOptions) {
            ArgumentNullException.ThrowIfNull(configureResizeOptions);

            VisioContainerOptions options = new();
            configureResizeOptions(options);
            return AddToContainer(containerId, memberIds, resizeToFit: true, resizeOptions: options);
        }

        /// <summary>
        /// Removes existing shapes from a Visio-native container.
        /// </summary>
        public VisioPage RemoveFromContainer(string containerId, IEnumerable<string> memberIds, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            VisioShape container = GetShape(containerId);
            this.RemoveFromContainer(container, ResolveShapes(memberIds, nameof(memberIds)), resizeToFit, resizeOptions);
            return this;
        }

        /// <summary>
        /// Removes existing shapes from a Visio-native container using inline resize option configuration.
        /// </summary>
        public VisioPage RemoveFromContainer(string containerId, IEnumerable<string> memberIds, Action<VisioContainerOptions> configureResizeOptions) {
            ArgumentNullException.ThrowIfNull(configureResizeOptions);

            VisioContainerOptions options = new();
            configureResizeOptions(options);
            return RemoveFromContainer(containerId, memberIds, resizeToFit: true, resizeOptions: options);
        }

        /// <summary>
        /// Gets typed metadata and membership information for a Visio-native container.
        /// </summary>
        public VisioContainerInfo ContainerInfo(string containerId) {
            return this.GetContainerInfo(GetShape(containerId));
        }

        /// <summary>
        /// Applies native container metadata and visual style to an existing Visio-native container.
        /// </summary>
        public VisioPage ApplyContainerOptions(string containerId, VisioContainerOptions options, bool refit = false) {
            this.ApplyContainerOptions(GetShape(containerId), options, refit);
            return this;
        }

        /// <summary>
        /// Updates native container metadata and visual style using a callback initialized from the current container state.
        /// </summary>
        public VisioPage ConfigureContainer(string containerId, Action<VisioContainerOptions> configureOptions, bool refit = false) {
            this.ConfigureContainer(GetShape(containerId), configureOptions, refit);
            return this;
        }

        /// <summary>
        /// Applies a reusable visual style to an existing Visio-native container.
        /// </summary>
        public VisioPage StyleContainer(string containerId, VisioShapeStyle style, bool refit = false) {
            ArgumentNullException.ThrowIfNull(style);

            return ConfigureContainer(containerId, options => options.ShapeStyle = style, refit);
        }

        /// <summary>
        /// Resizes a Visio-native container around its current members.
        /// </summary>
        public VisioPage RefitContainer(string containerId, VisioContainerOptions? options = null) {
            this.RefitContainer(GetShape(containerId), options);
            return this;
        }

        /// <summary>
        /// Resizes a Visio-native container around its current members using inline option configuration.
        /// </summary>
        public VisioPage RefitContainer(string containerId, Action<VisioContainerOptions> configureOptions) {
            ArgumentNullException.ThrowIfNull(configureOptions);

            VisioContainerOptions options = new();
            configureOptions(options);
            return RefitContainer(containerId, options);
        }

        private IReadOnlyList<VisioShape> ResolveShapes(IEnumerable<string> shapeIds, string parameterName) {
            ArgumentNullException.ThrowIfNull(shapeIds, parameterName);

            List<VisioShape> shapes = shapeIds.Select(GetShape).Distinct().ToList();
            if (shapes.Count == 0) {
                throw new ArgumentException("At least one shape id is required.", parameterName);
            }

            return shapes;
        }
    }
}
