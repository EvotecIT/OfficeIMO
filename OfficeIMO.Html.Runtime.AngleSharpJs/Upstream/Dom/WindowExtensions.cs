namespace AngleSharp.Js.Dom
{
    using AngleSharp;
    using AngleSharp.Attributes;
    using AngleSharp.Browser;
    using AngleSharp.Dom;
    using AngleSharp.Dom.Events;
    using AngleSharp.Html.Dom;
    using AngleSharp.Js.Attributes;
    using System;
    using System.Runtime.CompilerServices;

    /// <summary>
    /// Defines a set of extensions for the window object.
    /// </summary>
    [DomExposed("Window")]
    public static class WindowExtensions
    {
        private static readonly ConditionalWeakTable<IWindow, Console> Consoles =
            new ConditionalWeakTable<IWindow, Console>();
        private static readonly ConditionalWeakTable<IWindow, Worker> WorkerOwners =
            new ConditionalWeakTable<IWindow, Worker>();

        /// <summary>
        /// Posts a message.
        /// </summary>
        [DomName("postMessage")]
        public static void PostMessage(this IWindow window, String message, String targetOrigin = "*", Object transfer = null)
        {
            if (WorkerOwners.TryGetValue(window, out var owner))
            {
                owner.PostMessageToOwner(message);
                return;
            }

            var ev = new MessageEvent("message", false, false, message, targetOrigin);
            var document = window.Document;
            var loop = document.Context.GetService<IEventLoop>();
            loop.EnqueueAsync(_ => window.Fire(ev));
        }

        internal static void RegisterWorkerWindow(IWindow workerWindow, Worker owner)
        {
            if (workerWindow == null || owner == null)
            {
                return;
            }

            WorkerOwners.Remove(workerWindow);
            WorkerOwners.Add(workerWindow, owner);
        }

        /// <summary>
        /// Gets the parent window context.
        /// </summary>
        [DomName("parent")]
        [DomAccessor(Accessors.Getter)]
        public static IWindow Parent(this IWindow window)
        {
            var context = window.Document.Context;
            return GetWindow(context?.Parent) ?? window;
        }

        /// <summary>
        /// Gets the top window context.
        /// </summary>
        [DomName("top")]
        [DomAccessor(Accessors.Getter)]
        public static IWindow Top(this IWindow window)
        {
            var context = window.Document.Context;

            while (context?.Parent != null)
            {
                context = context.Parent;
            }

            return GetWindow(context) ?? window;
        }

        /// <summary>
        /// Gets the console instance. The same instance is returned for the same
        /// window, so that `window.console === window.console` holds and anything
        /// script puts on the console is still there on the next access.
        /// </summary>
        /// <param name="window"></param>
        /// <returns></returns>
        [DomName("console")]
        [DomAccessor(Accessors.Getter)]
        public static Console Console(this IWindow window)
        {
            return Consoles.GetValue(window, w => new Console(w));
        }

        /// <summary>
        /// Creates a new IHtmlImageElement instance.
        /// </summary>
        /// <param name="window"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        [DomConstructorFunction("Image")]
        public static IHtmlImageElement Image(this IWindow window, int? width = null, int? height = null)
        {
            var imageElement = window.Document.CreateElement(TagNames.Img) as IHtmlImageElement;

            if (width.HasValue)
            {
                imageElement.DisplayWidth = width.Value;
            }

            if (height.HasValue)
            {
                imageElement.DisplayHeight = height.Value;
            }

            return imageElement;
        }

        private static IWindow GetWindow(IBrowsingContext context) => context?.Active?.DefaultView;
    }
}
