using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using static envar.NativeMethods;

namespace envar
{
    internal class EnvironmentVariables
    {
        #region Constants

        public const string UserEnvironmentVariablesRegistryKeyPath = "Environment";

        public const string SystemEnvironmentVariablesRegistryKeyPath = "SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment";

        #endregion

        #region Nested Types

        [Flags]
        public enum SetVariableFlags
        {
            Append = 1,
            Overwrite = 2
        }

        #endregion

        #region Public Methods

        public static IDictionary<string, string> GetUserVariables()
        {
            return GetEnvironmentVariablesFromRegistry(EnvironmentVariableTarget.User);
        }

        public static IDictionary<string, string> GetSystemVariables()
        {
            return GetEnvironmentVariablesFromRegistry(EnvironmentVariableTarget.Machine);
        }

        public static IDictionary<string, string> GetProcessVariables(int processId)
        {
            return GetEnvironmentVariablesFromProcess(processId);
        }

        public static void SetUserVariable(string name, string value, SetVariableFlags flag)
        {
            SetEnvironmentVariable(name, value, flag, EnvironmentVariableTarget.User);
        }

        public static void SetSystemVariable(string name, string value, SetVariableFlags flag)
        {
            SetEnvironmentVariable(name, value, flag, EnvironmentVariableTarget.Machine);
        }

        #endregion

        #region Private Methods

        private static IDictionary<string, string> GetEnvironmentVariablesFromRegistry(EnvironmentVariableTarget target)
        {
            RegistryKey registryKey = null;
            switch (target)
            {
                case EnvironmentVariableTarget.Machine:
                    registryKey = Registry.LocalMachine.OpenSubKey(SystemEnvironmentVariablesRegistryKeyPath);
                    break;
                case EnvironmentVariableTarget.User:
                    registryKey = Registry.CurrentUser.OpenSubKey(UserEnvironmentVariablesRegistryKeyPath);
                    break;
            }

            if (registryKey == null)
            {
                throw new Exception("Failed to locate registry key.");
            }

            if (!(registryKey.ValueCount > 0))
            {
                return null;
            }

            var variables = new SortedDictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            foreach (var val in registryKey.GetValueNames())
            {
                var value = registryKey.GetValue(val)?.ToString();
                variables.Add(val, value);
            }

            return variables;
        }

        private static IDictionary<string, string> GetEnvironmentVariablesFromProcess(int processId)
        {
            var process = Process.GetProcessById(processId);
            var startInfo = process.StartInfo;

            return startInfo.EnvironmentVariables.ToSortedDictionary();
        }

        private static void SetEnvironmentVariable(string name, string value, SetVariableFlags flag,
            EnvironmentVariableTarget target)
        {
            var current = Environment.GetEnvironmentVariable(name, target);
            if (current != null)
            {
                if (flag.HasFlag(SetVariableFlags.Append))
                {
                    if (current.Split(';').Contains(value, true))
                        throw new ValueExistsException();
                    Environment.SetEnvironmentVariable(name, string.Join(";", current, value), target);
                }
                else if (flag.HasFlag(SetVariableFlags.Overwrite))
                {
                    Environment.SetEnvironmentVariable(name, value, target);
                }
                else
                {
                    throw new InvalidOperationException("The specified variable already exists.");
                }
            }
            else
            {
                Environment.SetEnvironmentVariable(name, value, target);
            }

            const int timeoutMilliseconds = 15000;

            var result = SendMessageTimeout(
                HWND_BROADCAST,
                WM_SETTINGCHANGE,
                UIntPtr.Zero,
                Marshal.StringToHGlobalUni("Environment"),
                SMTO_ABORTIFHUNG,
                timeoutMilliseconds,
                out _);

            if (result != IntPtr.Zero)
                return;

            var win32Err = Marshal.GetLastWin32Error();
            if ((ulong) win32Err == ERROR_TIMEOUT)
                throw new TimeoutException("The SendMessage operation timed out.");

            throw new Win32Exception(win32Err);
        }

        #endregion
    }

    internal class ValueExistsException : Exception
    {
        public ValueExistsException()
            :base("The value exists.")
        {            
        }
    }
}
