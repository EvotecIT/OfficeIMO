namespace AngleSharp.Js.Dom
{
    using AngleSharp.Attributes;
    using AngleSharp.Dom;
    using System;

    /// <summary>
    /// A set of stub implementations.
    /// </summary>
    [DomExposed("Element")]
    public static class ElementPseudoExtensions
    {
        /// <summary>
        /// Getter for the scrollLeft property.
        /// </summary>
        [DomName("scrollLeft")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetScrollLeft(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the scrollTop property.
        /// </summary>
        [DomName("scrollTop")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetScrollTop(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the scrollWidth property.
        /// </summary>
        [DomName("scrollWidth")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetScrollWidth(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the scrollHeight property.
        /// </summary>
        [DomName("scrollHeight")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetScrollHeight(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the clientLeft property.
        /// </summary>
        [DomName("clientLeft")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetClientLeft(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the clientTop property.
        /// </summary>
        [DomName("clientTop")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetClientTop(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the clientWidth property.
        /// </summary>
        [DomName("clientWidth")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetClientWidth(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the clientHeight property.
        /// </summary>
        [DomName("clientHeight")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetClientHeight(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the offsetLeft property.
        /// </summary>
        [DomName("offsetLeft")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetOffsetLeft(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the offsetTop property.
        /// </summary>
        [DomName("offsetTop")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetOffsetTop(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the offsetWidth property.
        /// </summary>
        [DomName("offsetWidth")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetOffsetWidth(this IElement element) => 0.0;

        /// <summary>
        /// Getter for the offsetHeight property.
        /// </summary>
        [DomName("offsetHeight")]
        [DomAccessor(Accessors.Getter)]
        public static Double GetOffsetHeight(this IElement element) => 0.0;

        /// <summary>
        /// Adds the focus in event.
        /// </summary>
        [DomName("focusin")]
        [DomAccessor(Accessors.Adder)]
        public static void AddFocusIn(this IElement element, DomEventHandler handler) =>
            element.AddEventListener("focusin", handler);

        /// <summary>
        /// Removes the focus in event.
        /// </summary>
        [DomName("focusin")]
        [DomAccessor(Accessors.Remover)]
        public static void RemoveFocusIn(this IElement element, DomEventHandler handler) =>
            element.RemoveEventListener("focusin", handler);

        /// <summary>
        /// Adds the focus out event.
        /// </summary>
        [DomName("focusout")]
        [DomAccessor(Accessors.Adder)]
        public static void AddFocusOut(this IElement element, DomEventHandler handler) =>
            element.AddEventListener("focusout", handler);

        /// <summary>
        /// Removes the focus out event.
        /// </summary>
        [DomName("focusout")]
        [DomAccessor(Accessors.Remover)]
        public static void RemoveFocusOut(this IElement element, DomEventHandler handler) =>
            element.RemoveEventListener("focusout", handler);

        /// <summary>
        /// Adds the unload event.
        /// </summary>
        [DomName("unload")]
        [DomAccessor(Accessors.Adder)]
        public static void AddUnload(this IElement element, DomEventHandler handler) =>
            element.AddEventListener("unload", handler);

        /// <summary>
        /// Removes the unload event.
        /// </summary>
        [DomName("unload")]
        [DomAccessor(Accessors.Remover)]
        public static void RemoveUnload(this IElement element, DomEventHandler handler) =>
            element.RemoveEventListener("unload", handler);

        /// <summary>
        /// Adds the contextmenu event.
        /// </summary>
        [DomName("contextmenu")]
        [DomAccessor(Accessors.Adder)]
        public static void AddContextMenu(this IElement element, DomEventHandler handler) =>
            element.AddEventListener("contextmenu", handler);

        /// <summary>
        /// Removes the contextmenu event.
        /// </summary>
        [DomName("contextmenu")]
        [DomAccessor(Accessors.Remover)]
        public static void RemoveContextMenu(this IElement element, DomEventHandler handler) =>
            element.RemoveEventListener("contextmenu", handler);
    }
}
