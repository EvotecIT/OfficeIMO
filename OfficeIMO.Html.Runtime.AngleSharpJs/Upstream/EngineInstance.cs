namespace AngleSharp.Js
{
    using AngleSharp.Dom;
    using AngleSharp.Io;
    using AngleSharp.Js.Cache;
    using AngleSharp.Text;
    using Jint;
    using Jint.Native;
    using Jint.Native.Json;
    using Jint.Native.Object;
    using Jint.Runtime.Interop;
    using System;
    using System.Collections.Generic;
    using System.Reflection;

    sealed class EngineInstance
    {
        #region Fields

        //  Jint's StackGuard.Disabled, which is internal.
        private const Int32 StackGuardDisabled = -1;

        private readonly Engine _engine;
        private readonly PrototypeCache _prototypes;
        private readonly ReferenceCache _references;
        private readonly LibrarySet _libs;
        private readonly DomNodeInstance _window;
        private readonly JsImportMap _importMap;

        #endregion

        #region ctor

        public EngineInstance(IWindow window, IDictionary<String, Object> assignments, IEnumerable<Assembly> libs, JsScriptingOptions options)
        {
            _importMap = new JsImportMap();

            _engine = new Engine((o) =>
            {
                o.EnableModules(new JsModuleLoader(this, window.Document, false));
                //  The handler answers out of the caches assigned right below, which only exist
                //  once this constructor returns. Jint wraps nothing while it is configuring
                //  itself, so that is safe - and Engine.Options is internal, so registering the
                //  handler afterwards is not an option.
                o.SetWrapObjectHandler(WrapObject);
                //  Left alone, the JS call stack is the native one, and a script recursing
                //  deeper than it holds takes the whole process down - a StackOverflowException
                //  cannot be caught. Guarded, the engine continues on a fresh stack and finally
                //  reports an ordinary "Maximum call stack size exceeded" error instead.
                o.Constraints.MaxExecutionStackCount = options.MaxCallStackDepth > 0 ? options.MaxCallStackDepth : StackGuardDisabled;
                options.ConfigureEngine?.Invoke(window, o);
            });
            _libs = new LibrarySet(libs);
            _prototypes = new PrototypeCache(_engine, _libs);
            _references = new ReferenceCache();

            foreach (var assignment in assignments)
            {
                _engine.SetValue(assignment.Key, assignment.Value);
            }

            _window = (DomNodeInstance)GetDomNode(window);

            foreach (var lib in libs)
            {
                this.AddConstructors(_window, lib);
                this.AddConstructorFunctions(_window, lib);
                this.AddInstances(_window, lib);
            }

            foreach (var property in _window.GetOwnProperties())
            {
                _engine.Global.FastSetProperty(property.Key.ToString(), property.Value);
            }

            foreach (var prototypeProperty in Window.Prototype.GetOwnProperties())
            {
                _engine.Global.FastSetProperty(prototypeProperty.Key.ToString(), prototypeProperty.Value);
            }

            _engine.Global.Prototype = _window.Prototype;
        }

        #endregion

        #region Properties

        public IEnumerable<Assembly> Libs => _libs.Assemblies;

        public DomNodeInstance Window => _window;

        public Engine Jint => _engine;

        public JsImportMap ImportMap => _importMap;

        #endregion

        #region Methods

        public ObjectInstance GetDomNode(Object obj) => _references.GetOrCreate(obj, CreateInstance);

        public ObjectInstance GetDomPrototype(Type type) => _prototypes.GetOrCreate(type, CreatePrototype);

        /// <summary>
        /// Gets the constructor object of the given type, building it on first ask. The
        /// prototype keeps it, so that naming the type and reading "constructor" off one of
        /// its instances arrive at the same object.
        /// </summary>
        public DomConstructorInstance GetDomConstructor(ConstructorDefinition definition)
        {
            //  Only the prototype of System.Object is not one of ours, and that type is not
            //  exposed as a constructor, so it never reaches this point.
            var prototype = GetDomPrototype(definition.Type);
            return DomPrototypeState.Of(prototype).GetConstructor(definition);
        }

        public JsValue RunScript(String source, String type, String sourceUrl)
        {
            if (string.IsNullOrEmpty(type))
            {
                type = MimeTypeNames.DefaultJavaScript;
            }

            lock (_engine)
            {
                if (MimeTypeNames.IsJavaScript(type))
                {
                    var prepared = ScriptCache.GetOrCreate(source);
                    //  An invalid result means the source did not parse; hand it to the
                    //  engine as text so the syntax error is reported as usual.
                    return prepared.IsValid ? _engine.Evaluate(prepared) : _engine.Evaluate(source);
                }
                else if (type.Isi("importmap"))
                {
                    return LoadImportMap(source);
                }
                else if (type.Isi("module"))
                {
                    // use a unique specifier to import the module into Jint
                    var specifier = sourceUrl ?? Guid.NewGuid().ToString();

                    return ImportModule(specifier, source);
                }
                else
                {
                    return JsValue.Undefined;
                }
            }
        }

        private JsValue LoadImportMap(String source)
        {
            //  The source is page content, so it must be handed to a JSON parser rather
            //  than pasted into a script: a single quote already breaks the parse, and
            //  anything after a closing quote would run as script.
            var importMap = new JsonParser(_engine).Parse(source).AsObject();

            if (importMap.TryGetValue("scopes", out var scopes))
            {
                var scopesObj = scopes.AsObject();

                foreach (var scopeProperty in scopesObj.GetOwnProperties())
                {
                    var scopePath = scopeProperty.Key.AsString();

                    if (_importMap.Scopes.ContainsKey(scopePath))
                    {
                        continue;
                    }

                    var scopeValue = new Dictionary<string, Uri>();

                    var scopeImports = scopesObj[scopePath].AsObject();

                    foreach (var scopeImportProperty in scopeImports.GetOwnProperties())
                    {
                        var scopeImportSpecifier = scopeImportProperty.Key.AsString();

                        if (!scopeValue.ContainsKey(scopeImportSpecifier))
                        {
                            scopeValue.Add(scopeImportSpecifier, new Uri(scopeImports[scopeImportSpecifier].AsString(), UriKind.RelativeOrAbsolute));
                        }
                    }

                    _importMap.Scopes.Add(scopePath, scopeValue);
                }
            }

            if (importMap.TryGetValue("imports", out var imports))
            {
                var importsObj = imports.AsObject();

                foreach (var importProperty in importsObj.GetOwnProperties())
                {
                    var importSpecifier = importProperty.Key.AsString();

                    if (!_importMap.Imports.ContainsKey(importSpecifier))
                    {
                        _importMap.Imports.Add(importSpecifier, new Uri(importsObj[importSpecifier].AsString(), UriKind.RelativeOrAbsolute));
                    }
                }
            }

            return JsValue.Undefined;
        }

        private JsValue ImportModule(String specifier, String source)
        {
            _engine.Modules.Add(specifier, source);
            _engine.Modules.Import(specifier);

            return JsValue.Undefined;
        }

        #endregion

        #region Helpers

        //  A collection is projected by the array-like proxy, which lets the engine read its
        //  indices and length directly; everything else by the ordinary one. The decision is a
        //  property of the type, so it is resolved once per type for the whole process.
        private ObjectInstance CreateInstance(Object obj)
        {
            var collection = obj.GetType().GetIndexedCollection();

            if (collection != null)
            {
                return new DomCollectionInstance(this, obj, collection);
            }

            return new DomNodeInstance(this, obj);
        }

        /// <summary>
        /// Builds this engine's prototype for a DOM type: an object over the member layout the
        /// whole process shares for that type, chained to the prototype of its base type and
        /// carrying what only this engine knows about it.
        /// </summary>
        private ObjectInstance CreatePrototype(Type type)
        {
            var shape = DomShapeCache.GetOrCreate(type, _libs, _engine);
            var prototype = shape.Shape.Instantiate(_engine, GetParentPrototype(type, shape.BaseType));
            JsObjectShape.SetHostState(prototype, new DomPrototypeState(this, prototype, shape));
            return prototype;
        }

        //  The base type may fold onto this very prototype - a class carrying no DOM name of its
        //  own shares the one of its nearest named ancestor. Asking for it would then re-enter
        //  the cache entry currently being built, so the fold is ruled out before the ask and the
        //  chain stops at Object.prototype, which is where it used to stop as well.
        private ObjectInstance GetParentPrototype(Type type, Type baseType) =>
            ReferenceEquals(_prototypes.Canonicalize(baseType), type)
                ? _engine.Intrinsics.Object.PrototypeObject
                : GetDomPrototype(baseType);

        /// <summary>
        /// Converts a value handed over from C#, which reaches Jint through JsValue.FromObject
        /// rather than through <see cref="EngineExtensions.ToJsValue"/>.
        /// </summary>
        /// <remarks>
        /// A DOM object has to arrive as the DOM proxy for the very same reason it does when the
        /// DOM itself yields one: the proxy is what carries the DOM members, and it is what makes
        /// the object script already holds and the one passed in from C# the same object. Anything
        /// else stays with Jint's CLR wrapper, which is what a host object is expected to be.
        /// </remarks>
        private ObjectInstance WrapObject(Engine engine, Object target, Type type) =>
            target.GetType().IsDomType() ? GetDomNode(target) : ObjectWrapper.Create(engine, target, type);

        #endregion
    }
}
