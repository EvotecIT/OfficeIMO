namespace AngleSharp
{
    using AngleSharp.Browser;
    using AngleSharp.Browser.Dom;
    using AngleSharp.Js;
    using AngleSharp.Js.Dom;
    using AngleSharp.Scripting;
    using System;

    /// <summary>
    /// Additional extensions for JavaScript scripting.
    /// </summary>
    public static class JsConfigurationExtensions
    {
        /// <summary>
        /// Includes a service to create a new console logger for the given context.
        /// </summary>
        /// <param name="configuration">The configuration to use.</param>
        /// <param name="createLogger">The delegate to create a new logger.</param>
        /// <returns>The new configuration.</returns>
        public static IConfiguration WithConsoleLogger(this IConfiguration configuration, Func<IBrowsingContext, IConsoleLogger> createLogger)
        {
            if (!configuration.Has<IConsoleLogger>())
            {
                configuration = configuration.With(createLogger);
            }

            return configuration;
        }

        /// <summary>
        /// Includes the thread-based JS event loop in the given context.
        /// </summary>
        /// <param name="configuration">The configuration to use.</param>
        /// <returns>The new configuration.</returns>
        public static IConfiguration WithEventLoop(this IConfiguration configuration) =>
            configuration.WithEventLoop(context => new JsEventLoop(context));

        /// <summary>
        /// Includes some event loop in the given context.
        /// </summary>
        /// <param name="configuration">The configuration to use.</param>
        /// <param name="loop">The existing loop to use.</param>
        /// <returns>The new configuration.</returns>
        public static IConfiguration WithEventLoop(this IConfiguration configuration, IEventLoop loop) =>
            configuration.WithOnly(loop);

        /// <summary>
        /// Includes some lazy initialized event loop in the given context.
        /// </summary>
        /// <param name="configuration">The configuration to use.</param>
        /// <param name="creator">The constructor function for the loop.</param>
        /// <returns>The new configuration.</returns>
        public static IConfiguration WithEventLoop(this IConfiguration configuration, Func<IBrowsingContext, IEventLoop> creator) =>
            configuration.WithOnly(creator);

        /// <summary>
        /// Sets scripting to true, registers the JavaScript engine and returns
        /// a new configuration with the scripting service and possible
        /// auxiliary services, if not yet registered.
        /// </summary>
        /// <param name="configuration">The configuration to use.</param>
        /// <returns>The new configuration.</returns>
        public static IConfiguration WithJs(this IConfiguration configuration) =>
            configuration.WithJs(new JsScriptingOptions());

        /// <summary>
        /// Sets scripting to true, registers the JavaScript engine with the
        /// given options and returns a new configuration with the scripting
        /// service and possible auxiliary services, if not yet registered.
        /// </summary>
        /// <param name="configuration">The configuration to use.</param>
        /// <param name="options">The options tuning the engine.</param>
        /// <returns>The new configuration.</returns>
        public static IConfiguration WithJs(this IConfiguration configuration, JsScriptingOptions options)
        {
            var service = new JsScriptingService(options);
            var observer = new EventAttributeObserver(service);
            var handler = new JsNavigationHandler(service);

            return configuration
                .WithOnly(observer)
                .With(handler)
                .With(service);
        }
    }
}
