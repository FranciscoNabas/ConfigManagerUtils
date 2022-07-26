using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Management;
using Microsoft.ConfigurationManagement.DesiredConfigurationManagement;
using System.Collections.Generic;
using System;
using System.Linq;
using System.IO;

#nullable enable
namespace ConfigManagerUtils.Utilities
{
    internal class UnmanagedCode
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr LoadLibrary(
            string lpLibFileName
        );
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool FreeLibrary(
            IntPtr hModule
        );
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(
            string lpModuleName
        );
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
    }
    public class Console
    {
        public class SearchResult
        {
            public string? Name { get; set; }
            public string? InstanceKey { get; set; }
            public string? Type { get; set; }
            public string? Path { get; set; }
            internal SearchResult(string name, string instanceKey, string type, string path) => (Name, InstanceKey, Type, Path) = (name, instanceKey, type, path);
            internal SearchResult() { }
        }

        public static List<SearchResult> SearchObject(string objectName, ObjectContainerType objectType, string siteServer)
        {
            ManagementScope scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS");
            scope.Connect();
            ObjectQuery? query = new ObjectQuery("Select Name From __NAMESPACE");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

            string? siteCode = GetSiteCode(searcher);
            scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS\\site_" + siteCode);
            scope.Connect();

            string queryExpression = "Select " + objectType.ClassPropertyName + ", " + objectType.ClassPropertyUniqueId +
                                    " From " + objectType.UnderlyingClassName +
                                    " Where " + objectType.ClassPropertyName + " Like '" + objectName + "'";
            query = new ObjectQuery(queryExpression);
            searcher = new ManagementObjectSearcher(scope, query);

            List<SearchResult> result = new List<SearchResult>();
            foreach (ManagementObject item in searcher.Get())
            {
                SearchResult single = new SearchResult();
                ObjectQuery objContQuery = new ObjectQuery("Select * From SMS_ObjectContainerItem Where InstanceKey = '" + item.Properties[objectType.ClassPropertyUniqueId].Value.ToString() + "'");
                ManagementObjectSearcher objContSearcher = new ManagementObjectSearcher(scope, objContQuery);

                var sGet = from ManagementObject x in objContSearcher.Get() select x;
                ManagementObject? objContResult = sGet.FirstOrDefault();

                single.Name = item.Properties[objectType.ClassPropertyName].Value.ToString();
                single.InstanceKey = item.Properties[objectType.ClassPropertyUniqueId].Value.ToString();
                single.Type = objectType.Name;
                if (objContResult is not null)
                {
                    single.Path = GetObjectFullPath(Convert.ToUInt32(objContResult.Properties["ContainerNodeID"].Value), scope);
                }

                result.Add(single);
            }

            return result;
        }

        public static List<SearchResult> SearchObject(string objectName, ObjectContainerType objectType, ManagementScope scope)
        {
            string queryExpression = "Select " + objectType.ClassPropertyName + ", " + objectType.ClassPropertyUniqueId +
                                                " From " + objectType.UnderlyingClassName +
                                                " Where " + objectType.ClassPropertyName + " Like '" + objectName + "'";
            ObjectQuery query = new ObjectQuery(queryExpression);
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

            List<SearchResult> result = new List<SearchResult>();
            foreach (ManagementObject item in searcher.Get())
            {
                SearchResult single = new SearchResult();
                ObjectQuery objContQuery = new ObjectQuery("Select * From SMS_ObjectContainerItem Where InstanceKey = '" + item.Properties[objectType.ClassPropertyUniqueId].Value.ToString() + "'");
                ManagementObjectSearcher objContSearcher = new ManagementObjectSearcher(scope, objContQuery);

                var sGet = from ManagementObject x in objContSearcher.Get() select x;
                ManagementObject? objContResult = sGet.FirstOrDefault();

                single.Name = item.Properties[objectType.ClassPropertyName].Value.ToString();
                single.InstanceKey = item.Properties[objectType.ClassPropertyUniqueId].Value.ToString();
                single.Type = objectType.Name;
                if (objContResult is not null)
                {
                    single.Path = GetObjectFullPath(Convert.ToUInt32(objContResult.Properties["ContainerNodeID"].Value), scope);
                }

                result.Add(single);
            }

            return result;
        }

        internal static string GetObjectFullPath(uint containerNodeId, ManagementScope scope)
        {
            SortedList<int, string> path = new SortedList<int, string>();
            int index = 0;
            ManagementObject? result = new ManagementObject();
            uint parentId = 1;
            while (parentId != 0)
            {
                ObjectQuery query = new ObjectQuery("Select Name, ParentContainerNodeID From SMS_ObjectContainerNode Where ContainerNodeID = '" + containerNodeId.ToString() + "'");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                var sGet = from ManagementObject x in searcher.Get() select x;
                result = sGet.FirstOrDefault();
                if (result is not null)
                {
                    string? name = result.Properties["Name"].Value.ToString();
                    if (!string.IsNullOrEmpty(name)) { path.Add(index, name); }
                    parentId = Convert.ToUInt32(result.Properties["ParentContainerNodeID"].Value);
                    containerNodeId = parentId;
                }
                index++;
            }
            string output = "";
            for (int i = path.Count - 1; i >= 0; i--) { output += @"\" + path[i]; }

            return Regex.Replace(output, "$", @"\");
        }
        internal static string? GetSiteCode(ManagementObjectSearcher searcher)
        {

            foreach (ManagementObject obj in searcher.Get())
            {
                string? smsNamespace = obj.Properties["Name"].Value.ToString();
                if (!string.IsNullOrEmpty(smsNamespace)) { obj.Dispose(); return smsNamespace.Replace("site_", ""); }
            }
            return null;
        }
    }
    public class Software
    {
        

        public static string GetFormatedWin32Error()
        {
            IntPtr errLib = UnmanagedCode.LoadLibrary("Ntdsbmsg.dll");
            int lastWin32Error = Marshal.GetLastWin32Error();
            StringBuilder buffer = new StringBuilder(512);

            int lResult = UnmanagedCode.FormatMessage(
                UnmanagedCode.FormatMessageFlags.FROM_SYSTEM |
                UnmanagedCode.FormatMessageFlags.IGNORE_INSERTS
                , IntPtr.Zero
                , Convert.ToUInt32(lastWin32Error)
                , 0
                , buffer
                , 512
                , IntPtr.Zero
            );

            UnmanagedCode.FreeLibrary(errLib);
            return buffer.ToString();
        }

        public static string GetFormatedWin32Error(int errorCode)
        {
            IntPtr errLib = UnmanagedCode.LoadLibrary("Ntdsbmsg.dll");
            StringBuilder buffer = new StringBuilder(512);

            int lResult = UnmanagedCode.FormatMessage(
                UnmanagedCode.FormatMessageFlags.FROM_SYSTEM |
                UnmanagedCode.FormatMessageFlags.IGNORE_INSERTS
                , IntPtr.Zero
                , Convert.ToUInt32(errorCode)
                , 0
                , buffer
                , 512
                , IntPtr.Zero
            );

            UnmanagedCode.FreeLibrary(errLib);
            return buffer.ToString();
        }

        public static string GetMsiProductCode(string filePath)
        {
            if (!File.Exists(filePath) || !filePath.EndsWith(".msi", StringComparison.Ordinal)) { throw new FileNotFoundException("Invalid file path. Make sure the file exists and it's a MSI."); }

            IntPtr database = new IntPtr();
            IntPtr view = new IntPtr();
            IntPtr record = new IntPtr();

            uint result = UnmanagedCode.MsiOpenDatabase(filePath, "MSIDBOPEN_READONLY", out database);
            result = UnmanagedCode.MsiDatabaseOpenView(database, "Select Property, Value From Property", out view);
            result = UnmanagedCode.MsiViewExecute(view, IntPtr.Zero);
            result = UnmanagedCode.MsiViewFetch(view, out record);

            uint GetBufferSize(IntPtr _record, uint index)
            {
                StringBuilder fetchBuff = new StringBuilder(1);
                uint bSize = 0;
                result = UnmanagedCode.MsiRecordGetString(record, 2, fetchBuff, out bSize);
                return bSize + 1;
            }


            string output = "";
            uint bSize = GetBufferSize(record, 1);
            StringBuilder buffer = new StringBuilder(Convert.ToInt32(bSize));

            while (record != IntPtr.Zero)
            {
                bSize = GetBufferSize(record, 1);
                buffer = new StringBuilder(Convert.ToInt32(bSize));
                result = UnmanagedCode.MsiRecordGetString(record, 1, buffer, out bSize);

                if (buffer.ToString() == "ProductCode")
                {
                    bSize = GetBufferSize(record, 2);
                    buffer = new StringBuilder(Convert.ToInt32(bSize));
                    result = UnmanagedCode.MsiRecordGetString(record, 2, buffer, out bSize);
                    output = buffer.ToString();
                    break;
                }

                result = UnmanagedCode.MsiViewFetch(view, out record);
            }

            result = UnmanagedCode.MsiCloseHandle(record);
            result = UnmanagedCode.MsiViewClose(view);
            result = UnmanagedCode.MsiCloseHandle(view);
            result = UnmanagedCode.MsiCloseHandle(database);

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