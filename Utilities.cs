using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Microsoft.ConfigurationManagement.DesiredConfigurationManagement;

namespace ConfigManagerUtils.Utilities
{
    public class Console
    {

    }
    public class Software
    {
        internal enum FormatMessageFlags : uint
        {
            ALLOCATE_BUFFER = 0x00000100,
            ARGUMENT_ARRAY = 0x00002000,
            FROM_HMODULE = 0x00000800,
            FROM_STRING = 0x00000400,
            FROM_SYSTEM = 0x00001000,
            IGNORE_INSERTS = 0x00000200
        }

        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        internal static extern uint MsiOpenDatabase(
            string szDatabasePath,
            string szPersist,
            out IntPtr phDatabase
        );

        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        internal static extern uint MsiDatabaseOpenView(
            IntPtr hDatabase,
            string szQuery,
            out IntPtr phView
        );

        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        internal static extern uint MsiViewExecute(
            IntPtr hView,
            IntPtr hRecord
        );

        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        internal static extern uint MsiViewFetch(
            IntPtr hView,
            out IntPtr hRecord
        );

        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        internal static extern uint MsiRecordGetString(
            IntPtr hRecord,
            uint iField,
            StringBuilder szValueBuf,
            out uint pcchValueBuf
        );

        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        internal static extern uint MsiViewClose(
            IntPtr hView
        );

        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        internal static extern uint MsiCloseHandle(
            IntPtr hAny
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int FormatMessage(
            FormatMessageFlags dwFlags,
            IntPtr lpSource,
            uint dwMessageId,
            uint dwLanguageId,
            StringBuilder msgOut,
            uint nSize,
            IntPtr Arguments
        );

        public static string GetFormatedWin32Error()
        {
            IntPtr errLib = new IntPtr();
            NativeLibrary.TryLoad("Ntdsbmsg.dll", out errLib);
            int lastWin32Error = Marshal.GetLastWin32Error();
            StringBuilder buffer = new StringBuilder(512);

            int lResult = FormatMessage(
                FormatMessageFlags.FROM_SYSTEM |
                FormatMessageFlags.IGNORE_INSERTS
                , IntPtr.Zero
                , Convert.ToUInt32(lastWin32Error)
                , 0
                , buffer
                , 512
                , IntPtr.Zero
            );

            NativeLibrary.Free(errLib);
            return buffer.ToString();
        }

        public static string GetFormatedWin32Error(int errorCode)
        {
            IntPtr errLib = new IntPtr();
            NativeLibrary.TryLoad("Ntdsbmsg.dll", out errLib);
            StringBuilder buffer = new StringBuilder(512);

            int lResult = FormatMessage(
                FormatMessageFlags.FROM_SYSTEM |
                FormatMessageFlags.IGNORE_INSERTS
                , IntPtr.Zero
                , Convert.ToUInt32(errorCode)
                , 0
                , buffer
                , 512
                , IntPtr.Zero
            );

            NativeLibrary.Free(errLib);
            return buffer.ToString();
        }

        public static string GetMsiProductCode(string filePath)
        {
            if (!File.Exists(filePath) || !filePath.EndsWith(".msi", StringComparison.Ordinal)) { throw new FileNotFoundException("Invalid file path. Make sure the file exists and it's a MSI."); }

            IntPtr database = new IntPtr();
            IntPtr view = new IntPtr();
            IntPtr record = new IntPtr();

            uint result = MsiOpenDatabase(filePath, "MSIDBOPEN_READONLY", out database);
            result = MsiDatabaseOpenView(database, "Select Property, Value From Property", out view);
            result = MsiViewExecute(view, IntPtr.Zero);
            result = MsiViewFetch(view, out record);

            uint GetBufferSize(IntPtr _record, uint index)
            {
                StringBuilder fetchBuff = new StringBuilder(1);
                uint bSize = 0;
                result = MsiRecordGetString(record, 2, fetchBuff, out bSize);
                return bSize + 1;
            }


            string output = "";
            uint bSize = GetBufferSize(record, 1);
            StringBuilder buffer = new StringBuilder(Convert.ToInt32(bSize));

            while (record != IntPtr.Zero)
            {
                bSize = GetBufferSize(record, 1);
                buffer = new StringBuilder(Convert.ToInt32(bSize));
                result = MsiRecordGetString(record, 1, buffer, out bSize);

                if (buffer.ToString() == "ProductCode")
                {
                    bSize = GetBufferSize(record, 2);
                    buffer = new StringBuilder(Convert.ToInt32(bSize));
                    result = MsiRecordGetString(record, 2, buffer, out bSize);
                    output = buffer.ToString();
                    break;
                }

                result = MsiViewFetch(view, out record);
            }

            result = MsiCloseHandle(record);
            result = MsiViewClose(view);
            result = MsiCloseHandle(view);
            result = MsiCloseHandle(database);

            return output;
        }

        public static bool ValidadeLocalPath(string localPath)
        {
            Uri uri = new Uri(localPath);
            if (uri.Scheme != "file" || uri.IsUnc == true || uri.Segments[1].Length != 3 || !Regex.IsMatch(uri.Segments[1], @"[A-z]:\/"))
            {
                throw new InvalidDataTypeException("Invalid object path. Enter a valid local path.");
            }
            if (localPath.Contains("/") || localPath.Contains(@"\\")) { throw new InvalidLogicalNameException("A folder or file name can't contain any of the following characters: \\ / : * ? \" < > |"); }
            string[] noChar = { ":", "*", "?", "\"", "<", ">", "|" };
            foreach (string wChar in noChar)
            {
                for (int i = 0; i < localPath.Length; i++)
                {
                    if (i > 3)
                    {
                        if (localPath[i].ToString() == wChar) { throw new InvalidLogicalNameException("A folder or file name can't contain any of the following characters: \\ / : * ? \" < > |"); }
                    }
                }
            }

            return true;
        }
    }
}