using System;
using System.Text;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using ConfigManagerUtils.Applications;
using ConfigManagerUtils;

namespace ConfigManagerUtils.StatusMessages
{
    public class Module
    {
        private int Id { get;}
        public String Name { get; }
        public static Module SMSClient => new(1, "climsgs.dll");
        public static Module SMSProvider => new(2, "provmsgs.dll");
        public static Module SMSServer => new(3, "srvmsgs.dll");
        private Module(int id, string name)
        {
            InternModule im = new InternModule(id, name);
            Id = im.Id;
            Name = im.Name;
        }
    }

    public enum Severity : uint
    {
        Information = 1073741824,
        Warning = 2147483648,
        Error = 3221225472
    }

    internal class UnmanagedCode
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(
            string lpModuleName
        );

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
            if (lModule != IntPtr.Zero) { NativeLibrary.Free(lModule); }

            IntPtr mHandle = new IntPtr();
            NativeLibrary.TryLoad(module.Name, out mHandle);
            if (mHandle == IntPtr.Zero) { throw new FileLoadException("Error loading library " + module.Name); }

            int bufferSize = 16384;
            StringBuilder output = new StringBuilder(bufferSize);

            int result = Utilities.FormatMessage(
                Utilities.FormatMessageFlags.FROM_HMODULE |
                Utilities.FormatMessageFlags.IGNORE_INSERTS
                , mHandle
                , (uint)severity | messageId
                , 0
                , output
                , Convert.ToUInt32(bufferSize)
                , IntPtr.Zero
            );

            if (result == 0)
            {
                string errMsg = Utilities.GetFormatedWin32Error();
                NativeLibrary.Free(mHandle);
                throw new SystemException(errMsg);
            }
            else
            {
                NativeLibrary.Free(mHandle);
                return output.ToString();
            }

        }

        public static string GetFormatedMessage(uint messageId, Severity severity, Module module, string[] insertString)
        {
            IntPtr lModule = UnmanagedCode.GetModuleHandle(module.Name);
            if (lModule != IntPtr.Zero) { NativeLibrary.Free(lModule); }

            IntPtr mHandle = new IntPtr();
            NativeLibrary.TryLoad(module.Name, out mHandle);
            if (mHandle == IntPtr.Zero) { throw new FileLoadException("Error loading library " + module.Name); }

            int bufferSize = 16384;
            StringBuilder output = new StringBuilder(bufferSize);

            int result = Utilities.FormatMessage(
                Utilities.FormatMessageFlags.FROM_HMODULE |
                Utilities.FormatMessageFlags.IGNORE_INSERTS
                , mHandle
                , (uint)severity | messageId
                , 0
                , output
                , Convert.ToUInt32(bufferSize)
                , IntPtr.Zero
            );

            if (result == 0)
            {
                string errMsg = Utilities.GetFormatedWin32Error();
                NativeLibrary.Free(mHandle);
                throw new SystemException(errMsg);
            }
            else
            {

                string message = Regex.Replace(output.ToString(), "%11|%12|%%3%4%5%6%7%8%9%10", "");
                for (int i = 0; i < insertString.Length; i++)
                {
                    message = message.Replace("%" + (i + 1).ToString(), insertString[i]);
                }
                NativeLibrary.Free(mHandle);
                return message;
            }

        }
        public static string GetFormatedMessage(StatusMessage _statusMessage)
        {
            IntPtr lModule = UnmanagedCode.GetModuleHandle(_statusMessage.Module.Name);
            if (lModule != IntPtr.Zero) { NativeLibrary.Free(lModule); }

            IntPtr mHandle = new IntPtr();
            NativeLibrary.TryLoad(_statusMessage.Module.Name, out mHandle);
            if (mHandle == IntPtr.Zero) { throw new FileLoadException("Error loading library " + _statusMessage.Module.Name); }

            int bufferSize = 16384;
            StringBuilder output = new StringBuilder(bufferSize);

            int result = Utilities.FormatMessage(
                Utilities.FormatMessageFlags.FROM_HMODULE |
                Utilities.FormatMessageFlags.IGNORE_INSERTS
                , mHandle
                , (uint)_statusMessage.Severity | _statusMessage.MessageId
                , 0
                , output
                , Convert.ToUInt32(bufferSize)
                , IntPtr.Zero
            );

            if (result == 0)
            {
                string errMsg = Utilities.GetFormatedWin32Error();
                NativeLibrary.Free(mHandle);
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
                    NativeLibrary.Free(mHandle);
                    return message;
                }
                else
                {
                    NativeLibrary.Free(mHandle);
                    return output.ToString();
                }

            }
        }
    }
}