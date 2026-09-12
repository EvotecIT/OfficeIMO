namespace AngleSharp.Js.Dom
{
    using AngleSharp.Attributes;
    using System;

    /// <summary>
    /// Represents information about the screen of the output device.
    /// </summary>
    [DomName("Screen")]
    [DomInstance("screen")]
    [DomExposed("Window")]
    public sealed class Screen
    {
        /// <summary>
        /// Gets the available width of the rendering surface of the output
        /// device, in CSS pixels.
        /// </summary>
        [DomName("availWidth")]
        public Int32 AvailableWidth => 1920;

        /// <summary>
        /// Gets the available height of the rendering surface of the output
        /// device, in CSS pixels.
        /// </summary>
        [DomName("availHeight")]
        public Int32 AvailableHeight => 1080;

        /// <summary>
        /// Gets the width of the output device, in CSS pixels.
        /// </summary>
        [DomName("width")]
        public Int32 Width => AvailableWidth;

        /// <summary>
        /// Gets the height of the output device, in CSS pixels.
        /// </summary>
        [DomName("height")]
        public Int32 Height => AvailableHeight;

        /// <summary>
        /// Gets the value 24 - by specification.
        /// </summary>
        [DomName("colorDepth")]
        public UInt32 ColorDepth => 24;

        /// <summary>
        /// Gets the value 24 - by specification.
        /// </summary>
        [DomName("pixelDepth")]
        public UInt32 PixelDepth => 24;
    }
}
