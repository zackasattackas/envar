using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Linq;

namespace envar
{
    internal static class ExtensionMethods
    {
        public static string Substring(this string s, Func<string, string> substringFunc)
        {
            return substringFunc?.Invoke(s);
        }

        public static IEnumerable<T> ToEnumerable<T>(this ICollection c)
        {
            var array = new T[c.Count];
            c.CopyTo(array, 0);

            return array;
        }

        public static SortedDictionary<string, string> ToSortedDictionary(this StringDictionary d)
        {
            var dictionary = new SortedDictionary<string, string>();
            foreach (DictionaryEntry entry in d)
            {
                dictionary.Add((string) entry.Key, (string) entry.Value);
            }
            return dictionary;
        }

        public static string GetProcessStatus(this Process p)
        {
            var thread = p.Threads[0];

            if (thread.ThreadState == ThreadState.Wait && thread.WaitReason == ThreadWaitReason.Suspended)
            {
                return ThreadWaitReason.Suspended.ToString();
            }

            return thread.ThreadState.ToString();
        }

        public static bool Contains(this IEnumerable<string> enumerable, string value, bool ignoreCase)
        {
            return enumerable.Any(a => a.Equals(value,
                ignoreCase ? StringComparison.CurrentCultureIgnoreCase : StringComparison.CurrentCulture));
        }
    }
}