namespace AngleSharp.Js.Cache
{
    using Jint.Native;
    using System;

    /// <summary>
    /// Everything the prototype of one DOM type is built from: the member layout the engine
    /// instantiates, the indexers the type offers, and what its constructor object is made of.
    /// </summary>
    /// <remarks>
    /// None of it belongs to an engine. A <see cref="JsObjectShape"/> is an immutable member
    /// layout that any number of engines may instantiate, the indexers are reflected accessors,
    /// and the constructor definition is reflection over a type - so the whole thing is resolved
    /// once per process and every document opened afterwards only allocates the object.
    /// </remarks>
    sealed class DomShape
    {
        public DomShape(Type type, Type baseType, String name, JsObjectShape shape, DomIndexers indexers, ConstructorDefinition constructor)
        {
            Type = type;
            BaseType = baseType;
            Name = name;
            Shape = shape;
            Indexers = indexers;
            Constructor = constructor;
        }

        /// <summary>
        /// Gets the type whose prototype this is - the class defining the DOM name, which many
        /// classes may share.
        /// </summary>
        public Type Type { get; }

        /// <summary>
        /// Gets the type the prototype above this one is built from.
        /// </summary>
        public Type BaseType { get; }

        /// <summary>
        /// Gets the DOM name of the type, or null if it carries none.
        /// </summary>
        public String Name { get; }

        /// <summary>
        /// Gets the member layout, shared by every engine's prototype for this type.
        /// </summary>
        public JsObjectShape Shape { get; }

        /// <summary>
        /// Gets the indexers the type offers, or null if it has none - which is the case for
        /// nearly every DOM type.
        /// </summary>
        public DomIndexers Indexers { get; }

        /// <summary>
        /// Gets what the constructor object of the type is built from, or null if the type is
        /// not exposed as one under any name the engine's libraries define.
        /// </summary>
        public ConstructorDefinition Constructor { get; }
    }
}
