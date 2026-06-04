using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        /// <summary>
        /// Adds a Visio-native container around existing shapes using inline option configuration.
        /// </summary>
        public VisioFluentPage Container(string id, string text, IEnumerable<string> memberIds, Action<VisioContainerOptions> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            VisioContainerOptions options = new();
            configure(options);
            return Container(id, text, memberIds, options);
        }

        /// <summary>
        /// Adds existing shapes to a Visio-native container.
        /// </summary>
        public VisioFluentPage AddToContainer(string containerId, IEnumerable<string> memberIds, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            VisioShape container = ResolveShape(containerId);
            Page.AddToContainer(container, ResolveShapes(memberIds, nameof(memberIds)), resizeToFit, resizeOptions);
            return this;
        }

        /// <summary>
        /// Adds existing shapes to a Visio-native container using inline resize option configuration.
        /// </summary>
        public VisioFluentPage AddToContainer(string containerId, IEnumerable<string> memberIds, Action<VisioContainerOptions> configureResizeOptions) {
            if (configureResizeOptions == null) {
                throw new ArgumentNullException(nameof(configureResizeOptions));
            }

            VisioContainerOptions options = new();
            configureResizeOptions(options);
            return AddToContainer(containerId, memberIds, resizeToFit: true, resizeOptions: options);
        }

        /// <summary>
        /// Removes existing shapes from a Visio-native container.
        /// </summary>
        public VisioFluentPage RemoveFromContainer(string containerId, IEnumerable<string> memberIds, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            VisioShape container = ResolveShape(containerId);
            Page.RemoveFromContainer(container, ResolveShapes(memberIds, nameof(memberIds)), resizeToFit, resizeOptions);
            return this;
        }

        /// <summary>
        /// Removes existing shapes from a Visio-native container using inline resize option configuration.
        /// </summary>
        public VisioFluentPage RemoveFromContainer(string containerId, IEnumerable<string> memberIds, Action<VisioContainerOptions> configureResizeOptions) {
            if (configureResizeOptions == null) {
                throw new ArgumentNullException(nameof(configureResizeOptions));
            }

            VisioContainerOptions options = new();
            configureResizeOptions(options);
            return RemoveFromContainer(containerId, memberIds, resizeToFit: true, resizeOptions: options);
        }

        /// <summary>
        /// Gets typed metadata and membership information for a Visio-native container.
        /// </summary>
        public VisioContainerInfo ContainerInfo(string containerId) {
            return Page.GetContainerInfo(ResolveShape(containerId));
        }

        /// <summary>
        /// Applies native container metadata and visual style to an existing Visio-native container.
        /// </summary>
        public VisioFluentPage ApplyContainerOptions(string containerId, VisioContainerOptions options, bool refit = false) {
            Page.ApplyContainerOptions(ResolveShape(containerId), options, refit);
            return this;
        }

        /// <summary>
        /// Updates native container metadata and visual style using a callback initialized from the current container state.
        /// </summary>
        public VisioFluentPage ConfigureContainer(string containerId, Action<VisioContainerOptions> configureOptions, bool refit = false) {
            Page.ConfigureContainer(ResolveShape(containerId), configureOptions, refit);
            return this;
        }

        /// <summary>
        /// Applies a reusable visual style to an existing Visio-native container.
        /// </summary>
        public VisioFluentPage StyleContainer(string containerId, VisioShapeStyle style, bool refit = false) {
            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            return ConfigureContainer(containerId, options => options.ShapeStyle = style, refit);
        }

        /// <summary>
        /// Resizes a Visio-native container around its current members.
        /// </summary>
        public VisioFluentPage RefitContainer(string containerId, VisioContainerOptions? options = null) {
            Page.RefitContainer(ResolveShape(containerId), options);
            return this;
        }

        /// <summary>
        /// Resizes a Visio-native container around its current members using inline option configuration.
        /// </summary>
        public VisioFluentPage RefitContainer(string containerId, Action<VisioContainerOptions> configureOptions) {
            if (configureOptions == null) {
                throw new ArgumentNullException(nameof(configureOptions));
            }

            VisioContainerOptions options = new();
            configureOptions(options);
            return RefitContainer(containerId, options);
        }

        private IReadOnlyList<VisioShape> ResolveShapes(IEnumerable<string> shapeIds, string parameterName) {
            if (shapeIds == null) {
                throw new ArgumentNullException(parameterName);
            }

            List<VisioShape> shapes = shapeIds.Select(ResolveShape).Distinct().ToList();
            if (shapes.Count == 0) {
                throw new ArgumentException("At least one shape id is required.", parameterName);
            }

            return shapes;
        }
    }
}
