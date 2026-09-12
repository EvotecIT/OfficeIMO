namespace AngleSharp.Js.Proxies
{
    using Jint.Native;
    using Jint.Native.Object;
    using Jint.Runtime;
    using Jint.Runtime.Descriptors;
    using System.Reflection;

    sealed class DomConstructorFunctionInstance : Constructor
    {
        private readonly EngineInstance _instance;
        private readonly MethodInfo _constructorFunction;

        public DomConstructorFunctionInstance(EngineInstance instance, MethodInfo constructorFunction, string name) : base(instance.Jint, name)
        {
            _instance = instance;
            _constructorFunction = constructorFunction;

            //  Jint's Constructor leaves the prototype at Object.prototype, which would make a
            //  DOM constructor the one function in the engine without call, apply or bind.
            Prototype = (ObjectInstance)instance.Jint.Intrinsics.Function.Get("prototype");

            //  Image is a second way of naming HTMLImageElement rather than a type of its own,
            //  so it publishes the prototype of what it returns. Without one, "instanceof"
            //  against it is a TypeError rather than an answer.
            SetOwnProperty("prototype", new PropertyDescriptor(
                instance.GetDomPrototype(constructorFunction.ReturnType), false, false, false));
        }

        public override ObjectInstance Construct(JsValue[] arguments, JsValue newTarget)
        {
            try
            {
                return _instance.Call(_constructorFunction, _instance.Window, arguments) as ObjectInstance;
            }
            catch
            {
                throw new JavaScriptException(_instance.Jint.Intrinsics.Error);
            }
        }
    }
}
