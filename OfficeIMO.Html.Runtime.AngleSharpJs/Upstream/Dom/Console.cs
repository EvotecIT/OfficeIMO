using AngleSharp.Attributes;
using AngleSharp.Dom;

namespace AngleSharp.Js.Dom
{
    /// <summary>
    /// Defines the console.
    /// </summary>
    [DomName("Console")]
    public sealed class Console
    {
        private readonly IConsoleLogger _logger;

        /// <summary>
        /// Creates a new console.
        /// </summary>
        /// <param name="window"></param>
        [DomConstructor]
        public Console(IWindow window)
        {
            var context = window.Document.Context;
            _logger = context.GetService<IConsoleLogger>();
        }

        /// <summary>
        /// Outputs a message to the console.
        /// </summary>
        /// <param name="objs"></param>
        [DomName("log")]
        public void Log(params object[] objs)
        {
            _logger?.Log(objs);
        }
    }
}
