namespace AngleSharp.Js.Cache
{
    using AngleSharp.Attributes;
    using AngleSharp.Dom;
    using System;
    using System.Collections.Concurrent;
    using System.Reflection;

    /// <summary>
    /// Describes what building the arguments of a method needs to know about it.
    /// All of it is fixed by the method itself, so it is worked out once instead of
    /// on every call from script.
    /// </summary>
    sealed class MethodDescription
    {
        private static readonly ConcurrentDictionary<MethodBase, MethodDescription> _descriptions =
            new ConcurrentDictionary<MethodBase, MethodDescription>();

        private MethodDescription(MethodBase method)
        {
            var parameters = method.GetParameters();
            var descriptions = new ParameterDescription[parameters.Length];

            for (var i = 0; i < parameters.Length; i++)
            {
                descriptions[i] = new ParameterDescription(parameters[i]);
            }

            Parameters = descriptions;
            InitDict = method.GetCustomAttribute<DomInitDictAttribute>();
            TakesWindow = parameters.Length > 0 && parameters[0].ParameterType == typeof(IWindow);
            TakesParamArray = parameters.Length > 0 &&
                parameters[parameters.Length - 1].GetCustomAttribute<ParamArrayAttribute>() != null;
        }

        public static MethodDescription Of(MethodBase method) =>
            _descriptions.GetOrAdd(method, m => new MethodDescription(m));

        public ParameterDescription[] Parameters { get; }

        public DomInitDictAttribute InitDict { get; }

        public Boolean TakesWindow { get; }

        public Boolean TakesParamArray { get; }
    }

    /// <summary>
    /// The parts of a <see cref="ParameterInfo"/> that are read when arguments are
    /// built. Reading them from reflection means walking metadata every time.
    /// </summary>
    readonly struct ParameterDescription
    {
        public ParameterDescription(ParameterInfo parameter)
        {
            Name = parameter.Name;
            ParameterType = parameter.ParameterType;
            IsOptional = parameter.IsOptional;
            DefaultValue = parameter.IsOptional ? parameter.DefaultValue : null;
        }

        public String Name { get; }

        public Type ParameterType { get; }

        public Boolean IsOptional { get; }

        public Object DefaultValue { get; }
    }
}
