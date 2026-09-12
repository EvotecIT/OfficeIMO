namespace AngleSharp.Js
{
    using AngleSharp.Js.Cache;
    using Jint.Native;
    using Jint.Native.Object;
    using Jint.Native.Symbol;
    using Jint.Runtime;
    using Jint.Runtime.Descriptors;
    using Jint.Runtime.Interop;
    using System;
    using System.Reflection;

    sealed class DomConstructorInstance : Constructor
    {
        private readonly ConstructorInfo _constructor;
        private readonly EngineInstance _instance;
        private readonly ObjectInstance _objectPrototype;
        private readonly Type _type;

        public DomConstructorInstance(EngineInstance engine, ConstructorDefinition definition)
            : base(engine.Jint, definition.Name)
        {
            var toString = new ClrFunction(Engine, "toString", ToString);
            _objectPrototype = engine.GetDomPrototype(definition.Type);
            _instance = engine;
            _constructor = definition.Info;
            _type = definition.Type;

            //  Jint's Constructor leaves the prototype at Object.prototype, which would make a
            //  DOM constructor the one function in the engine without call, apply or bind.
            Prototype = (ObjectInstance)engine.Jint.Intrinsics.Function.Get("prototype");

            FastSetProperty("toString", new PropertyDescriptor(toString, true, false, true));

            //  Nothing is written back onto the prototype. Its member layout declares
            //  "constructor" as a slot resolved on first read, and that read arrives here, so the
            //  two directions still meet at one object without this end having to reach over.
            SetOwnProperty("prototype", new PropertyDescriptor(_objectPrototype, false, false, false));
        }

        /// <summary>
        /// Answers "instanceof" itself, because the prototype chain cannot always carry the
        /// answer: a mixin such as ParentNode has no class of its own to hang a prototype off,
        /// and the closed instantiations of IHtmlCollection&lt;T&gt; are separate types that a
        /// single prototype cannot stand for. Built on first ask - most types are never asked.
        /// </summary>
        public override PropertyDescriptor GetOwnProperty(JsValue property)
        {
            if (property == GlobalSymbolRegistry.HasInstance)
            {
                var descriptor = base.GetOwnProperty(property);

                if (descriptor == PropertyDescriptor.Undefined)
                {
                    var hasInstance = new ClrFunction(Engine, "[Symbol.hasInstance]", HasInstance, 1, PropertyFlag.Configurable);
                    descriptor = new PropertyDescriptor(hasInstance, false, false, false);
                    SetOwnProperty(property, descriptor);
                }

                return descriptor;
            }

            return base.GetOwnProperty(property);
        }

        private JsValue HasInstance(JsValue thisObject, JsValue[] arguments)
        {
            var value = arguments.Length > 0 ? arguments[0] : JsValue.Undefined;

            if (value is IDomProxy node && IsInstance(node.Value))
            {
                return JsBoolean.True;
            }

            //  Not a DOM object of this type, but something may still have been given this
            //  prototype - Object.create(HTMLDivElement.prototype) is an instance in a browser.
            return InheritsFromPrototype(value as ObjectInstance) ? JsBoolean.True : JsBoolean.False;
        }

        private Boolean IsInstance(Object value)
        {
            var type = value.GetType();

            if (_type.IsAssignableFrom(type))
            {
                return true;
            }

            //  IHtmlCollection<T> is exposed as HTMLCollection, so an instance of any of its
            //  closed forms answers to it.
            if (_type.GetTypeInfo().IsGenericTypeDefinition)
            {
                foreach (var contract in type.GetTypeInfo().ImplementedInterfaces)
                {
                    if (contract.GetTypeInfo().IsGenericType && contract.GetGenericTypeDefinition() == _type)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private Boolean InheritsFromPrototype(ObjectInstance obj)
        {
            while (obj != null)
            {
                obj = obj.Prototype;

                if (ReferenceEquals(obj, _objectPrototype))
                {
                    return true;
                }
            }

            return false;
        }

        public override ObjectInstance Construct(JsValue[] arguments, JsValue newTarget)
        {
            if (_constructor == null)
            {
                throw new JavaScriptException("Illegal constructor.");
            }

            try
            {
                var parameters = _instance.BuildArgs(_constructor, arguments);
                var obj = _constructor.Invoke(parameters);
                return _instance.GetDomNode(obj);
            }
            catch
            {
                throw new JavaScriptException(_instance.Jint.Intrinsics.Error);
            }
        }

        protected override JsValue Call(JsValue thisObject, JsValue[] arguments)
        {
            if (_constructor != null)
            {
                throw new JavaScriptException("Only call the constructor with the new keyword.");
            }

            return Construct(arguments, null);
        }

        private JsValue ToString(JsValue thisObj, JsValue[] arguments) => ToString();
    }
}
