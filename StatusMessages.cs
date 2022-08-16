using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using ConfigManagerUtils.Utilities;
using System.Reflection;
using System;

#nullable enable
namespace ConfigManagerUtils.StatusMessages
{
    public enum Severity : uint
    {
        Information = 1073741824,
        Warning = 2147483648,
        Error = 3221225472
    }

    public class StatusMessage
    {
        public Module Module { get; set; }
        public uint MessageId { get; set; }
        public Severity Severity { get; set; }
        public string[]? InsertString { get; set; }

        public StatusMessage(Module module, uint messageId, Severity severity)
        {
            Module = module;
            MessageId = messageId;
            Severity = severity;
        }

        public StatusMessage(Module module, uint messageId, Severity severity, string[] insertString)
        {
            Module = module;
            MessageId = messageId;
            Severity = severity;
            InsertString = insertString;
        }

        public static string GetFormatedMessage(uint messageId, Severity severity, Module module)
        {
            IntPtr lModule = UnmanagedCode.GetModuleHandle(module.Name);
            if (lModule != IntPtr.Zero) { UnmanagedCode.FreeLibrary(lModule); }

            string modulePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\UnmanagedLibraries\\" + module.Name;
            IntPtr mHandle = UnmanagedCode.LoadLibrary(modulePath);
            if (mHandle == IntPtr.Zero) {
                string errMsg = Software.GetFormatedWin32Error();
                throw new SystemException(module.Name + ": " + errMsg);
            }

            int bufferSize = 16384;
            StringBuilder output = new StringBuilder(bufferSize);

            int result = UnmanagedCode.FormatMessage(
                UnmanagedCode.FormatMessageFlags.FROM_HMODULE |
                UnmanagedCode.FormatMessageFlags.IGNORE_INSERTS
                , mHandle
                , (uint)severity | messageId
                , 0
                , output
                , Convert.ToUInt32(bufferSize)
                , IntPtr.Zero
            );

            if (result == 0)
            {
                string errMsg = Software.GetFormatedWin32Error();
                UnmanagedCode.FreeLibrary(mHandle);
                throw new SystemException(errMsg);
            }
            else
            {
                UnmanagedCode.FreeLibrary(mHandle);
                return output.ToString();
            }

        }

        public static string GetFormatedMessage(uint messageId, Severity severity, Module module, string[] insertString)
        {
            IntPtr lModule = UnmanagedCode.GetModuleHandle(module.Name);
            if (lModule != IntPtr.Zero) { UnmanagedCode.FreeLibrary(lModule); }

            string modulePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\UnmanagedLibraries\\" + module.Name;
            IntPtr mHandle = UnmanagedCode.LoadLibrary(modulePath);
            if (mHandle == IntPtr.Zero) { throw new FileLoadException("Error loading library " + module.Name); }

            int bufferSize = 16384;
            StringBuilder output = new StringBuilder(bufferSize);

            int result = UnmanagedCode.FormatMessage(
                UnmanagedCode.FormatMessageFlags.FROM_HMODULE |
                UnmanagedCode.FormatMessageFlags.IGNORE_INSERTS
                , mHandle
                , (uint)severity | messageId
                , 0
                , output
                , Convert.ToUInt32(bufferSize)
                , IntPtr.Zero
            );

            if (result == 0)
            {
                string errMsg = Software.GetFormatedWin32Error();
                UnmanagedCode.FreeLibrary(mHandle);
                throw new SystemException(errMsg);
            }
            else
            {

                string message = Regex.Replace(output.ToString(), "%11|%12|%%3%4%5%6%7%8%9%10", "");
                for (int i = 0; i < insertString.Length; i++)
                {
                    message = message.Replace("%" + (i + 1).ToString(), insertString[i]);
                }
                UnmanagedCode.FreeLibrary(mHandle);
                return message;
            }

        }
        public static string GetFormatedMessage(StatusMessage _statusMessage)
        {
            IntPtr lModule = UnmanagedCode.GetModuleHandle(_statusMessage.Module.Name);
            if (lModule != IntPtr.Zero) { UnmanagedCode.FreeLibrary(lModule); }

            string modulePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\UnmanagedLibraries\\" + _statusMessage.Module.Name;
            IntPtr mHandle = UnmanagedCode.LoadLibrary(modulePath);
            if (mHandle == IntPtr.Zero) { throw new FileLoadException("Error loading library " + _statusMessage.Module.Name); }

            int bufferSize = 16384;
            StringBuilder output = new StringBuilder(bufferSize);

            int result = UnmanagedCode.FormatMessage(
                UnmanagedCode.FormatMessageFlags.FROM_HMODULE |
                UnmanagedCode.FormatMessageFlags.IGNORE_INSERTS
                , mHandle
                , (uint)_statusMessage.Severity | _statusMessage.MessageId
                , 0
                , output
                , Convert.ToUInt32(bufferSize)
                , IntPtr.Zero
            );

            if (result == 0)
            {
                string errMsg = Software.GetFormatedWin32Error();
                UnmanagedCode.FreeLibrary(mHandle);
                throw new SystemException(errMsg);
            }
            else
            {
                if (_statusMessage.InsertString != null)
                {
                    string message = Regex.Replace(output.ToString(), "%11|%12|%%3%4%5%6%7%8%9%10", "");
                    for (int i = 0; i < _statusMessage.InsertString.Length; i++)
                    {
                        message = message.Replace("%" + (i + 1).ToString(), _statusMessage.InsertString[i]);
                    }
                    UnmanagedCode.FreeLibrary(mHandle);
                    return message;
                }
                else
                {
                    UnmanagedCode.FreeLibrary(mHandle);
                    return output.ToString();
                }

            }
        }
    }
}