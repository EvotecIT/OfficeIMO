namespace AngleSharp.Html.Dom.Events
{
    using System;

    /// <summary>
    /// A couple of useful extensions for the modifier list.
    /// </summary>
    static class ModifierExtensions
    {
        public static Boolean IsCtrlPressed(this String modifierList)
        {
            return modifierList.ContainsKey("Control");
        }

        public static Boolean IsMetaPressed(this String modifierList)
        {
            return modifierList.ContainsKey("Meta");
        }

        public static Boolean IsShiftPressed(this String modifierList)
        {
            return modifierList.ContainsKey("Shift");
        }

        public static Boolean IsAltPressed(this String modifierList)
        {
            return modifierList.ContainsKey("Alt");
        }

        public static Boolean ContainsKey(this String modifierList, String key)
        {
            if (String.IsNullOrWhiteSpace(modifierList) || String.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            foreach (var modifier in modifierList.Split((Char[]?)null, StringSplitOptions.RemoveEmptyEntries))
            {
                if (String.Equals(modifier, key, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
