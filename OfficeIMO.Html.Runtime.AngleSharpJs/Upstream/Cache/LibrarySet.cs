namespace AngleSharp.Js.Cache
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;

    /// <summary>
    /// The AngleSharp libraries an engine exposes to script, as a value.
    /// </summary>
    /// <remarks>
    /// What a DOM type looks like from script is not decided by the type alone: which extension
    /// members it gains and which name its constructor is found under both depend on the set of
    /// libraries the browsing context registered services from, and on the order they are
    /// consulted in. Anything cached for the whole process therefore has to be keyed by this as
    /// well as by the type - hence the value semantics.
    /// </remarks>
    sealed class LibrarySet : IEquatable<LibrarySet>
    {
        private readonly Assembly[] _assemblies;
        private readonly Int32 _hash;

        public LibrarySet(IEnumerable<Assembly> assemblies)
        {
            var list = new List<Assembly>();

            foreach (var assembly in assemblies)
            {
                list.Add(assembly);
            }

            _assemblies = list.ToArray();
            _hash = _assemblies.Length;

            for (var i = 0; i < _assemblies.Length; i++)
            {
                _hash = (_hash * 397) ^ _assemblies[i].GetHashCode();
            }
        }

        public IReadOnlyList<Assembly> Assemblies => _assemblies;

        public Boolean Equals(LibrarySet other)
        {
            if (ReferenceEquals(this, other))
            {
                return true;
            }

            if (other == null || other._assemblies.Length != _assemblies.Length)
            {
                return false;
            }

            for (var i = 0; i < _assemblies.Length; i++)
            {
                if (!ReferenceEquals(_assemblies[i], other._assemblies[i]))
                {
                    return false;
                }
            }

            return true;
        }

        public override Boolean Equals(Object obj) => Equals(obj as LibrarySet);

        public override Int32 GetHashCode() => _hash;
    }
}
