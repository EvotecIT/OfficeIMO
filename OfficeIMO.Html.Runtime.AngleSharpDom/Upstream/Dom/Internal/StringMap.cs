namespace AngleSharp.Dom
{
    using AngleSharp.Text;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Represents a list of DOMTokens.
    /// </summary>
    sealed class StringMap : IStringMap
    {
        #region Fields

        private readonly String _prefix;
        private readonly Element _parent;

        #endregion

        #region ctor

        internal StringMap(String prefix, Element parent)
        {
            _prefix = prefix;
            _parent = parent;
        }

        #endregion

        #region Index

        public String? this[String name]
        {
            get => HasForbiddenDash(name) ? null : _parent.GetOwnAttribute(_prefix + ToAttributeName(name));
            set
            {
                if (HasForbiddenDash(name))
                {
                    throw new DomException(DomError.Syntax);
                }

                var attributeName = _prefix + ToAttributeName(name);
                if (value == null)
                {
                    _parent.RemoveAttribute(attributeName);
                }
                else
                {
                    _parent.SetAttribute(attributeName, value);
                }
            }
        }

        #endregion

        #region Methods

        public void Remove(String name)
        {
            _parent.RemoveAttribute(_prefix + ToAttributeName(name));
        }

        public Boolean Contains(String name)
        {
            return !HasForbiddenDash(name) && _parent.HasOwnAttribute(_prefix + ToAttributeName(name));
        }

        #endregion

        #region Helper

        private static Boolean HasForbiddenDash(String name)
        {
            for (var i = 0; i < name.Length - 1; i++)
            {
                if (name[i] == '-' && name[i + 1].IsLowercaseAscii())
                {
                    return true;
                }
            }

            return false;
        }

        private static String ToAttributeName(String name)
        {
            var builder = new System.Text.StringBuilder(name.Length);
            foreach (var character in name)
            {
                if (character.IsUppercaseAscii())
                {
                    builder.Append('-');
                    builder.Append(Char.ToLowerInvariant(character));
                }
                else
                {
                    builder.Append(character);
                }
            }

            return builder.ToString();
        }

        private static String ToPropertyName(String name)
        {
            var builder = new System.Text.StringBuilder(name.Length);
            for (var i = 0; i < name.Length; i++)
            {
                if (name[i] == '-' && i + 1 < name.Length && name[i + 1].IsLowercaseAscii())
                {
                    builder.Append(Char.ToUpperInvariant(name[++i]));
                }
                else
                {
                    builder.Append(name[i]);
                }
            }

            return builder.ToString();
        }

        #endregion

        #region IEnumerable Implementation

        public IEnumerator<KeyValuePair<String, String>> GetEnumerator()
        {
            foreach (var attr in _parent.Attributes)
            {
                if (attr.NamespaceUri is null && attr.Name.StartsWith(_prefix, StringComparison.Ordinal) &&
                    !attr.Name.Substring(_prefix.Length).Any(character => character.IsUppercaseAscii()))
                {
                    var name = attr.Name.Remove(0, _prefix.Length);
                    var value = attr.Value;
                    yield return new KeyValuePair<String, String>(ToPropertyName(name), value);
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion
    }
}
