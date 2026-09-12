namespace AngleSharp.Js
{
    using AngleSharp.Dom;
    using AngleSharp.Io;
    using AngleSharp.Text;
    using Jint;
    using Jint.Runtime.Modules;
    using System.IO;
    using System;
    using System.Linq;

    internal class JsModuleLoader : DefaultModuleLoader
    {
        private readonly EngineInstance _instance;
        private readonly IResourceLoader _resourceLoader;
        private readonly IElement _scriptElement;
        private readonly string _documentUrl;

        public JsModuleLoader(EngineInstance instance, IDocument document, bool restrictToBasePath = true) : base (document.Url, restrictToBasePath)
        {
            _instance = instance;
            _resourceLoader = document.Context.GetService<IResourceLoader>();
            _scriptElement = document.CreateElement(TagNames.Script);
            _documentUrl = document.Url;
        }

        public override ResolvedSpecifier Resolve(string referencingModuleLocation, ModuleRequest moduleRequest)
        {
            if (referencingModuleLocation != null && _instance.ImportMap.Scopes.Count > 0)
            {
                foreach (var scopePath in _instance.ImportMap.Scopes.Keys.OrderByDescending(k => k.Length))
                {
                    if (referencingModuleLocation.Contains(scopePath))
                    {
                        var scopeImports = _instance.ImportMap.Scopes[scopePath];

                        if (scopeImports.TryGetValue(moduleRequest.Specifier, out var scopeModuleUrl))
                        {
                            if (!scopeModuleUrl.IsAbsoluteUri)
                            {
                                scopeModuleUrl = new Uri(new Uri(_documentUrl), scopeModuleUrl);
                            }

                            return new ResolvedSpecifier(
                                    moduleRequest,
                                    moduleRequest.Specifier,
                                    scopeModuleUrl,
                                    SpecifierType.RelativeOrAbsolute);
                        }
                    }
                }
            }

            if (_instance.ImportMap.Imports.TryGetValue(moduleRequest.Specifier, out var moduleUrl))
            {
                if (!moduleUrl.IsAbsoluteUri)
                {
                    moduleUrl = new Uri(new Uri(_documentUrl), moduleUrl);
                }

                return new ResolvedSpecifier(
                        moduleRequest,
                        moduleRequest.Specifier,
                        moduleUrl,
                        SpecifierType.RelativeOrAbsolute);
            }

            // Before passing to the base default module loader, make sure any relative paths are prepended with a dot
            // as Jint will otherwise resolve them as absolute file paths when running on Linux
            if (moduleRequest.Specifier.StartsWith("/"))
            {
                moduleRequest = new ModuleRequest($".{moduleRequest.Specifier}", moduleRequest.Attributes);
            }

            if (referencingModuleLocation?.StartsWith("/") == true)
            {
                referencingModuleLocation = $".{referencingModuleLocation}";
            }
            
            return base.Resolve(referencingModuleLocation, moduleRequest);
        }

        protected override string LoadModuleContents(Engine engine, ResolvedSpecifier resolved)
        {
            if (resolved.Uri?.Scheme.IsOneOf(ProtocolNames.Http, ProtocolNames.Https) == true)
            {
                return FetchModule(resolved.Uri);
            }
            else
            {
                return base.LoadModuleContents(engine, resolved);
            }
        }

        private string FetchModule(Uri moduleUrl)
        {
            if (_resourceLoader == null)
            {
                return string.Empty;
            }

            var importUrl = Url.Convert(moduleUrl);

            var request = new ResourceRequest(_scriptElement, importUrl);

            var response = _resourceLoader.FetchAsync(request).Task.Result;

            string content;

            using (var streamReader = new StreamReader(response.Content))
            {
                content = streamReader.ReadToEnd();
            }

            return content;
        }
    }
}
