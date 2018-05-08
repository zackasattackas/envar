using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using static envar.NativeMethods;

namespace envar
{
    internal class Program
    {
        private static void Main(string[] args)
        { 
            var copyArgs = args;
            PrintCopyright();

            NormalizeArgs(copyArgs);

            if (!copyArgs.Any() || copyArgs[0] == "-?")
            {
                PrintHelp(true);
                return;
            }

            #region Main App Code

            try
            {
                var action = copyArgs[0];
                switch (action)
                {
                    case "-l" :
                        {
                            var target = GetEnvironmentVariableTarget(copyArgs.Length == 1 ? "-u" : copyArgs[1]);

                            IDictionary<string, string> envVariables;
                            switch (target)
                            {
                                case EnvironmentVariableTarget.Machine :
                                    PrintRegistryPath(target.Value);
                                    envVariables = EnvironmentVariables.GetSystemVariables();
                                    break;
                                case EnvironmentVariableTarget.User :
                                    PrintRegistryPath(target.Value);
                                    envVariables = EnvironmentVariables.GetUserVariables();
                                    break;
                                case EnvironmentVariableTarget.Process :
                                    if (!(copyArgs.Length > 2))
                                    {
                                        throw new ArgumentException("To view the variables for a process, a process identifier (PID) must be supplied.");
                                    }
                                    var pid = int.Parse(copyArgs[2]);
                                    PrintProcess(Process.GetProcessById(pid));
                                    envVariables = EnvironmentVariables.GetProcessVariables(pid);
                                    break;
                                case null:
                                default :
                                    throw new ArgumentOutOfRangeException();
                            }
                            PrintDictionaryEntries(envVariables);
                        }
                        break;
                    case "-s" :
                        {
                            var targetStr = copyArgs.Length == 1 ? "-u" : copyArgs[1];
                            var target = GetEnvironmentVariableTarget(targetStr);

                            if (target == EnvironmentVariableTarget.Process)
                            {
                                throw new InvalidOperationException("A process' environment block cannot be modified except by a parent process.");
                            }

                            string varName = null;
                            string varValue = null;
                            EnvironmentVariables.SetVariableFlags? flag = null;
                            for (var i = 2; i < copyArgs.Length; i++)
                            {
                                if (char.IsLetter(copyArgs[i][0]))
                                {
                                    continue;
                                }
                                switch (copyArgs[i])
                                {
                                    case "-a" :
                                        if (flag != null)
                                        {
                                            throw new ArgumentException("The -a and -o options are mutually exclusive.");
                                        }
                                        flag = EnvironmentVariables.SetVariableFlags.Append;
                                        break;
                                    case "-o" :
                                        if (flag != null)
                                        {
                                            throw new ArgumentException("The -a and -o options are mutually exclusive.");
                                        }
                                        flag = EnvironmentVariables.SetVariableFlags.Overwrite;
                                        break;
                                    case "-n" :
                                        if (copyArgs.Length < i + 1)
                                        {
                                            throw new ArgumentException("A variable name must be specified after the -n option.");
                                        }
                                        varName = copyArgs[i + 1];
                                        break;
                                    case "-v" :
                                        if (copyArgs.Length < i + 1)
                                        {
                                            throw new ArgumentException("A variable value must be specified after the -v option.");
                                        }
                                        varValue = copyArgs[i + 1];
                                        break;
                                    default :
                                        throw new ArgumentException($"Unexpected argument: {copyArgs[i]}");
                                }
                            }

                            if (flag == null)
                            {
                                flag = EnvironmentVariables.SetVariableFlags.Append;
                            }

                            switch (target)
                            {
                                case EnvironmentVariableTarget.User :
                                    EnvironmentVariables.SetUserVariable(varName, varValue, flag.Value);
                                    break;
                                case EnvironmentVariableTarget.Machine :
                                    EnvironmentVariables.SetSystemVariable(varName, varValue, flag.Value);
                                    break;
                                case EnvironmentVariableTarget.Process:
                                case null:
                                default :
                                    throw new ArgumentException("Invalid Environment target. Valid options are -m (Machine) and -u (User).");
                            }
                        }
                        break;
                    case "-b":
                        {
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
                    default :
                        throw new ArgumentOutOfRangeException();
                }
            }
            catch (Exception e)
            {
                string message = null;
                if (e is IndexOutOfRangeException)
                {
                    message = "One or more command line options were not provided.";
                }
                Console.WriteLine($"\r\nERROR: {message ?? e.Message}");
                //PrintHelp();
            }

            #endregion
#if DEBUG
            if (!Debugger.IsAttached)
            {
                return;
            }
            Console.Write("\r\nPress any key to exit...");
            Console.ReadKey();
#endif 
        }

        #region Console Output

        private static void PrintCopyright()
        {
            Console.WriteLine("\r\nEnvar - Environment Variables Utility");
            Console.WriteLine("Zack Bolin (2017)");
        }
        private static void PrintDictionaryEntries(IDictionary<string, string> values)
        {
            var width = values.Keys.Max(k => k.Length);
            var format = $"{{0,-{width}}} : {{1}}";
            foreach (var value in values)
            {
                var keyValue = value.Key == "Path" || value.Key == "PSModulePath"
                    ? string.Join("; ", value.Value.Split(';'))
                    : value.Value;
                Console.WriteLine(format, value.Key, GetWrappedText(width + 3, keyValue));
            }
        }
        private static void PrintUsage() 
        {
            Console.WriteLine(@"    
  Usage: envar {
                    -l [[-m] | [-u] | [-p <process id>]]
               }
               {
                    -s [[-m] | [-u] [[-a] | [-o]]] -n <name> -v <value | NULL>
               }");
            Console.WriteLine();
        }
        private static void PrintHelp(bool printExamples = false)
        {
            PrintUsage();
            const string format = "     {0,-5} {1}";
            const int left = 11;
            Console.WriteLine("  LIST");
            Console.WriteLine(format, "-l", GetWrappedText(left, "List the environment variables. If no environment target is specified, the user's environment block is used."));
            Console.WriteLine(format, "-p", GetWrappedText(left, "If specified, an attempt is made to locate a process with the specified PID and enumerate the variables available in the process' environment block."));
            Console.WriteLine(format, string.Empty, GetWrappedText(left, "Note: Because the Windows kernel reuses these ID's once no longer in use, it is possible that the original process has exited and a new process or thread has been assigned the same ID."));
            Console.WriteLine("\r\n  MODIFY");
            Console.WriteLine(format, "-s", GetWrappedText(left, "Create or update an environment variable using the specified name and value. If -a or -o is not specified, and the variable already exists, an error will be returned."));
            Console.WriteLine(format, "-a", GetWrappedText(left, "Append the specified value with a semi-colon delimiter if the variable already exists. This option is the default."));
            Console.WriteLine(format, "-o", GetWrappedText(left, "Overwrite the value of the existing variable."));
            Console.WriteLine("\r\n  GLOBAL");
            Console.WriteLine(format, "-m", GetWrappedText(left, "The variables accessed will be the System values found in the HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\Session Manager\\Environment registry key."));
            Console.WriteLine(format, "-u", GetWrappedText(left, "The variables access will be the User values found in the HKEY_CURRENT_USER\\Environment registry key."));
            Console.WriteLine(format, "-b", GetWrappedText(left, "Broadcasts a WM_SETTINGCHANGE to all windows. This forces applications to update their Environment blocks without the need for a reboot."));
            Console.WriteLine(format, "-?", "Display this help information.");

            if (!printExamples)
            {
                return;
            }

            var response = ReadKeyWithPrompt("\r\n::::Press <Spacebar> to see usage examples::::", ConsoleKey.Spacebar);

            if (!response)
            {
                return;
            }
            PrintExamples();
        }
        private static void PrintProcess(Process p)
        {
            const string format = "\r\nName: {0, -15} Pid: {1,5} Status: {2,-10} File: {3}\r\n";
            Console.WriteLine(format, p.ProcessName, p.Id, p.GetProcessStatus(), p.StartInfo.FileName);
        }
        private static void PrintRegistryPath(EnvironmentVariableTarget target)
        {
            switch (target)
            {
                case EnvironmentVariableTarget.Machine :
                    Console.WriteLine($"\r\nPath: HKEY_LOCAL_MACHINE\\{EnvironmentVariables.SystemEnvironmentVariablesRegistryKeyPath}\r\n");
                    break;
                case EnvironmentVariableTarget.User :
                    Console.WriteLine($"\r\nPath: HKEY_CURRENT_USER\\{EnvironmentVariables.UserEnvironmentVariablesRegistryKeyPath}\r\n");
                    break;
            }
        }
        private static void PrintExamples()
        {
            const string cli = "C:\\>envar.exe";
            Console.WriteLine("  EXAMPLES");
            Console.WriteLine("    [1] List the system environment variables.");
            Console.WriteLine($"\t{cli} -l -m");
            Console.WriteLine("\r\n    [2] List the environment for the process with ID 15222.");
            Console.WriteLine($"\t{cli} -l -p 15222");
            Console.WriteLine("\r\n    [3] Append a directory the current user's \"Path\" variable.");
            Console.WriteLine($"\t{cli} -s -u -a -n Path -v \"C:\\Program Files\\Sysinternals\"");
            Console.WriteLine("\r\n    [4] Delete the \"PSModulePath\" variable for the current user.");
            Console.WriteLine($"\t{cli} -s -u -o -n PSModulePath -v NULL");
        }

        #endregion

        #region Helpers

        private static void NormalizeArgs(IList<string> args)
        {
            var sb = new StringBuilder();
            for (var i = 0; i < args.Count; i++)
            {
                if (char.IsLetter(args[i][0]))
                {
                    continue;
                }
                for (var j = 0; j < args[i].Length; j++)
                {
                    if (j == 0 && args[i][j] == '/')
                    {
                        sb.Append('-');
                    }
                    else
                    {
                        sb.Append(char.ToLower(args[i][j]));
                    }                    
                }

                args[i] = sb.ToString();
                sb.Clear();
            }
        }
        private static string GetWrappedText(int startPosition, string s)
        {
            if (startPosition + s.Length < Console.BufferWidth)
            {
                return s;
            }
            var builder = new StringBuilder();
            var chunk = s
                .Substring(0, Console.BufferWidth - startPosition)
                .Substring(sub => sub.Substring(0, sub.Contains(" ") ? sub.LastIndexOf(" ", StringComparison.Ordinal) : sub.Length));
            builder.AppendLine(chunk.Trim());
            builder.Append(new string(' ', startPosition));
            builder.Append(GetWrappedText(startPosition, s.Substring(chunk.Length).Trim()));
            return builder.ToString();
        }
        private static EnvironmentVariableTarget? GetEnvironmentVariableTarget(string s)
        {
            switch (s)
            {
                case "-m":
                    return EnvironmentVariableTarget.Machine;
                case "-u":
                    return EnvironmentVariableTarget.User;
                case "-p":
                    return EnvironmentVariableTarget.Process;
                default:
                    throw new ArgumentException("Invalid Environment target. Valid options are -m (Machine), -u (User) or -p (Process).");
            }
        }
        private static void ClearConsoleLine(int cursorTop)
        {
            Console.SetCursorPosition(0, cursorTop);
            Console.Write(new string(' ', Console.BufferWidth));
            Console.SetCursorPosition(0, Console.CursorTop - 1);
        }
        private static bool ReadKeyWithPrompt(string prompt, ConsoleKey key, bool clearPrompt = true)
        {
            Console.Write(prompt);
            var keyPressed = Console.ReadKey().Key;
            if (clearPrompt)
            {
                ClearConsoleLine(Console.CursorTop);
            }
            return keyPressed == key;
        }

        #endregion
    }
}
