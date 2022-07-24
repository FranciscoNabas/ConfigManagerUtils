using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Management;
using Microsoft.ConfigurationManagement.ApplicationManagement;
using Microsoft.ConfigurationManagement.ApplicationManagement.Serialization;
using Microsoft.ConfigurationManagement.DesiredConfigurationManagement;
using Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions;
using Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules;
using ConfigManagerUtils.Applications.DeploymentTypes;

namespace ConfigManagerUtils.Applications
{
    public class Utilities
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

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class Application : Microsoft.ConfigurationManagement.ApplicationManagement.Application
    {
        public Application(ObjectId id) : base(id) { new Microsoft.ConfigurationManagement.ApplicationManagement.Application(id); }

        public static Application CreateMsiApplication(string authoringScope, string applicationName, Msi deploymentType)
        {
            ObjectId id = new ObjectId("Scope_Id" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            app.Version = 1;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("Scope_Id" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            DeploymentType dt = new DeploymentType(id, MsiInstaller.TechnologyId);
            dt.Title = "DT_" + applicationName + "_MSI";
            dt.Version = 1;

            Content content = new Content();
            ContentFile file = new ContentFile();
            ContentRef contentRef = new ContentRef();
            content.Location = deploymentType.ContentPath;
            file.Name = deploymentType.InstallerName;
            contentRef.Id = content.Id;
            content.Files.Add(file);

            MsiInstaller? installer = dt.Installer as MsiInstaller;
            if (installer is not null)
            {
                installer.ProductCode = deploymentType.ProductCode;
                installer.InstallCommandLine = deploymentType.InstallString;
                installer.UninstallCommandLine = deploymentType.UninstallString;
                installer.ExecutionContext = deploymentType.Context;
                installer.UserInteractionMode = deploymentType.UserInteractionMode;
                installer.Contents.Add(content);
                installer.InstallContent = contentRef;
            }

            app.DeploymentTypes.Add(dt);

            return app;
        }

        public static void CreateAndPublishMsiApplication(string authoringScope, string applicationName, Msi deploymentType, string siteServer)
        {
            ObjectId id = new ObjectId("Scope_Id" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            app.Version = 1;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("Scope_Id" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            DeploymentType dt = new DeploymentType(id, MsiInstaller.TechnologyId);
            dt.Title = "DT_" + applicationName + "_MSI";
            dt.Version = 1;

            Content content = new Content();
            ContentFile file = new ContentFile();
            ContentRef contentRef = new ContentRef();
            content.Location = deploymentType.ContentPath;
            file.Name = deploymentType.InstallerName;
            contentRef.Id = content.Id;
            content.Files.Add(file);

            MsiInstaller? installer = dt.Installer as MsiInstaller;
            if (installer is not null)
            {
                installer.ProductCode = deploymentType.ProductCode;
                installer.InstallCommandLine = deploymentType.InstallString;
                installer.UninstallCommandLine = deploymentType.UninstallString;
                installer.ExecutionContext = deploymentType.Context;
                installer.UserInteractionMode = deploymentType.UserInteractionMode;
                installer.Contents.Add(content);
                installer.InstallContent = contentRef;
            }

            app.DeploymentTypes.Add(dt);

            PublishApplication(siteServer, app);
        }

        public static void CreateAndPublishMsiApplication(string authoringScope, string applicationName, Msi deploymentType, string siteServer, string targetPath)
        {
            ObjectId id = new ObjectId("Scope_Id" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            app.Version = 1;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("Scope_Id" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            DeploymentType dt = new DeploymentType(id, MsiInstaller.TechnologyId);
            dt.Title = "DT_" + applicationName + "_MSI";
            dt.Version = 1;

            Content content = new Content();
            ContentFile file = new ContentFile();
            ContentRef contentRef = new ContentRef();
            content.Location = deploymentType.ContentPath;
            file.Name = deploymentType.InstallerName;
            contentRef.Id = content.Id;
            content.Files.Add(file);

            MsiInstaller? installer = dt.Installer as MsiInstaller;
            if (installer is not null)
            {
                installer.ProductCode = deploymentType.ProductCode;
                installer.InstallCommandLine = deploymentType.InstallString;
                installer.UninstallCommandLine = deploymentType.UninstallString;
                installer.ExecutionContext = deploymentType.Context;
                installer.UserInteractionMode = deploymentType.UserInteractionMode;
                installer.Contents.Add(content);
                installer.InstallContent = contentRef;
            }

            app.DeploymentTypes.Add(dt);

            PublishApplication(siteServer, app, targetPath);
        }

        public static Application CreateGeneralApplication(string authoringScope, string applicationName, DeploymentTypes.Script deploymentType)
        {
            ObjectId id = new ObjectId("ScopeId_" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            app.Version = 1;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("ScopeId_" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            DeploymentType dt = new DeploymentType(id, ScriptInstaller.TechnologyId);
            dt.Title = "DT_" + applicationName + "_Script";
            dt.Version = 1;

            ScriptInstaller? installer = dt.Installer as ScriptInstaller;
            if (installer is not null)
            {
                Microsoft.ConfigurationManagement.ApplicationManagement.Script detectionScript = new Microsoft.ConfigurationManagement.ApplicationManagement.Script();
                detectionScript.RunAs32Bit = deploymentType.RunAs32Bit;
                detectionScript.Language = deploymentType.Language.ToString();
                detectionScript.Text = deploymentType.ScriptBlock.ToString();

                installer.DetectionScript = detectionScript;
                installer.InstallCommandLine = deploymentType.InstallString;
                installer.UninstallCommandLine = deploymentType.UninstallString;
                installer.ExecutionContext = deploymentType.Context;
                installer.UserInteractionMode = deploymentType.UserInteractionMode;

                if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                {
                    Content content = new Content();
                    ContentFile file = new ContentFile();
                    ContentRef contentRef = new ContentRef();
                    content.Location = deploymentType.ContentPath;
                    file.Name = deploymentType.InstallerName;
                    contentRef.Id = content.Id;
                    content.Files.Add(file);
                    installer.Contents.Add(content);
                    installer.InstallContent = contentRef;
                }

            }

            app.DeploymentTypes.Add(dt);

            return app;
        }

        public static Application CreateGeneralApplication(string authoringScope, string applicationName, DeploymentTypes.Enhanced deploymentType)
        {
            ObjectId id = new ObjectId("ScopeId_" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            int appVersion = 1;
            app.Version = appVersion;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("ScopeId_" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            if (!string.IsNullOrEmpty(deploymentType.InstallerName) && deploymentType.InstallerName.EndsWith(".msi"))
            {
                DeploymentType dt = new DeploymentType(id, MsiInstaller.TechnologyId);
                dt.Title = "DT_" + applicationName + "_MSI";
                dt.Version = 1;

                MsiInstaller? installer = dt.Installer as MsiInstaller;
                if (installer is not null)
                {
                    GetEnhancedDetectionMethod(app, deploymentType, ref installer);
                    if (string.IsNullOrEmpty(deploymentType.InstallString)) { installer.InstallCommandLine = "msiexec.exe /I \"" + deploymentType.InstallerName + "\""; }
                    else { installer.InstallCommandLine = deploymentType.InstallString; }
                    installer.ExecutionContext = deploymentType.Context;
                    installer.UserInteractionMode = deploymentType.UserInteractionMode;

                    if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                    {
                        Content content = new Content();
                        ContentFile file = new ContentFile();
                        ContentRef contentRef = new ContentRef();
                        content.Location = deploymentType.ContentPath;
                        file.Name = deploymentType.InstallerName;
                        contentRef.Id = content.Id;
                        content.Files.Add(file);
                        installer.Contents.Add(content);
                        installer.InstallContent = contentRef;

                        if (string.IsNullOrEmpty(deploymentType.UninstallString))
                        {
                            try
                            {
                                string productCode = Utilities.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
                                installer.UninstallCommandLine = "msiexec.exe /X \"" + productCode + "\"";
                            }
                            catch { installer.UninstallCommandLine = null; }
                        }
                    }
                }

                app.DeploymentTypes.Add(dt);
            }
            else
            {
                DeploymentType dt = new DeploymentType(id, ScriptInstaller.TechnologyId);
                dt.Title = "DT_" + applicationName + "_Script";
                dt.Version = 1;

                ScriptInstaller? installer = dt.Installer as ScriptInstaller;
                if (installer is not null)
                {
                    GetEnhancedDetectionMethod(app, deploymentType, ref installer);
                    if (string.IsNullOrEmpty(deploymentType.InstallString)) { installer.InstallCommandLine = "msiexec.exe /I \"" + deploymentType.InstallerName + "\""; }
                    else { installer.InstallCommandLine = deploymentType.InstallString; }
                    installer.ExecutionContext = deploymentType.Context;
                    installer.UserInteractionMode = deploymentType.UserInteractionMode;

                    if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                    {
                        Content content = new Content();
                        ContentFile file = new ContentFile();
                        ContentRef contentRef = new ContentRef();
                        content.Location = deploymentType.ContentPath;
                        file.Name = deploymentType.InstallerName;
                        contentRef.Id = content.Id;
                        content.Files.Add(file);
                        installer.Contents.Add(content);
                        installer.InstallContent = contentRef;

                        if (string.IsNullOrEmpty(deploymentType.UninstallString))
                        {
                            try
                            {
                                string productCode = Utilities.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
                                installer.UninstallCommandLine = "msiexec.exe /X \"" + productCode + "\"";
                            }
                            catch { installer.UninstallCommandLine = null; }
                        }
                    }
                }

                app.DeploymentTypes.Add(dt);
            }

            return app;
        }

        public static void CreateAndPublishGeneralApplication(string authoringScope, string applicationName, DeploymentTypes.Script deploymentType, string siteServer)
        {
            ObjectId id = new ObjectId("ScopeId_" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            app.Version = 1;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("ScopeId_" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            DeploymentType dt = new DeploymentType(id, ScriptInstaller.TechnologyId);
            dt.Title = "DT_" + applicationName + "_Script";
            dt.Version = 1;

            ScriptInstaller? installer = dt.Installer as ScriptInstaller;
            if (installer is not null)
            {
                Microsoft.ConfigurationManagement.ApplicationManagement.Script detectionScript = new Microsoft.ConfigurationManagement.ApplicationManagement.Script();
                detectionScript.RunAs32Bit = deploymentType.RunAs32Bit;
                detectionScript.Language = deploymentType.Language.ToString();
                detectionScript.Text = deploymentType.ScriptBlock.ToString();

                installer.DetectionScript = detectionScript;
                installer.InstallCommandLine = deploymentType.InstallString;
                installer.UninstallCommandLine = deploymentType.UninstallString;
                installer.ExecutionContext = deploymentType.Context;
                installer.UserInteractionMode = deploymentType.UserInteractionMode;

                if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                {
                    Content content = new Content();
                    ContentFile file = new ContentFile();
                    ContentRef contentRef = new ContentRef();
                    content.Location = deploymentType.ContentPath;
                    file.Name = deploymentType.InstallerName;
                    contentRef.Id = content.Id;
                    content.Files.Add(file);
                    installer.Contents.Add(content);
                    installer.InstallContent = contentRef;
                }

            }

            app.DeploymentTypes.Add(dt);
            PublishApplication(siteServer, app);
        }

        public static void CreateAndPublishGeneralApplication(string authoringScope, string applicationName, DeploymentTypes.Script deploymentType, string siteServer, string targetPath)
        {
            ObjectId id = new ObjectId("ScopeId_" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            app.Version = 1;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("ScopeId_" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            DeploymentType dt = new DeploymentType(id, ScriptInstaller.TechnologyId);
            dt.Title = "DT_" + applicationName + "_Script";
            dt.Version = 1;

            ScriptInstaller? installer = dt.Installer as ScriptInstaller;
            if (installer is not null)
            {
                Microsoft.ConfigurationManagement.ApplicationManagement.Script detectionScript = new Microsoft.ConfigurationManagement.ApplicationManagement.Script();
                detectionScript.RunAs32Bit = deploymentType.RunAs32Bit;
                detectionScript.Language = deploymentType.Language.ToString();
                detectionScript.Text = deploymentType.ScriptBlock.ToString();

                installer.DetectionScript = detectionScript;
                installer.InstallCommandLine = deploymentType.InstallString;
                installer.UninstallCommandLine = deploymentType.UninstallString;
                installer.ExecutionContext = deploymentType.Context;
                installer.UserInteractionMode = deploymentType.UserInteractionMode;

                if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                {
                    Content content = new Content();
                    ContentFile file = new ContentFile();
                    ContentRef contentRef = new ContentRef();
                    content.Location = deploymentType.ContentPath;
                    file.Name = deploymentType.InstallerName;
                    contentRef.Id = content.Id;
                    content.Files.Add(file);
                    installer.Contents.Add(content);
                    installer.InstallContent = contentRef;
                }

            }

            app.DeploymentTypes.Add(dt);
            PublishApplication(siteServer, app, targetPath);
        }

        public static void CreateAndPublishGeneralApplication(string authoringScope, string applicationName, DeploymentTypes.Enhanced deploymentType, string siteServer)
        {
            ObjectId id = new ObjectId("ScopeId_" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            int appVersion = 1;
            app.Version = appVersion;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("ScopeId_" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            if (!string.IsNullOrEmpty(deploymentType.InstallerName) && deploymentType.InstallerName.EndsWith(".msi"))
            {
                DeploymentType dt = new DeploymentType(id, MsiInstaller.TechnologyId);
                dt.Title = "DT_" + applicationName + "_MSI";
                dt.Version = 1;

                MsiInstaller? installer = dt.Installer as MsiInstaller;
                if (installer is not null)
                {
                    GetEnhancedDetectionMethod(app, deploymentType, ref installer);
                    if (string.IsNullOrEmpty(deploymentType.InstallString)) { installer.InstallCommandLine = "msiexec.exe /I \"" + deploymentType.InstallerName + "\""; }
                    else { installer.InstallCommandLine = deploymentType.InstallString; }
                    installer.ExecutionContext = deploymentType.Context;
                    installer.UserInteractionMode = deploymentType.UserInteractionMode;

                    if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                    {
                        Content content = new Content();
                        ContentFile file = new ContentFile();
                        ContentRef contentRef = new ContentRef();
                        content.Location = deploymentType.ContentPath;
                        file.Name = deploymentType.InstallerName;
                        contentRef.Id = content.Id;
                        content.Files.Add(file);
                        installer.Contents.Add(content);
                        installer.InstallContent = contentRef;

                        if (string.IsNullOrEmpty(deploymentType.UninstallString))
                        {
                            try
                            {
                                string productCode = Utilities.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
                                installer.UninstallCommandLine = "msiexec.exe /X \"" + productCode + "\"";
                            }
                            catch { installer.UninstallCommandLine = null; }
                        }
                    }
                }

                app.DeploymentTypes.Add(dt);
            }
            else
            {
                DeploymentType dt = new DeploymentType(id, ScriptInstaller.TechnologyId);
                dt.Title = "DT_" + applicationName + "_Script";
                dt.Version = 1;

                ScriptInstaller? installer = dt.Installer as ScriptInstaller;
                if (installer is not null)
                {
                    GetEnhancedDetectionMethod(app, deploymentType, ref installer);
                    if (string.IsNullOrEmpty(deploymentType.InstallString)) { installer.InstallCommandLine = "msiexec.exe /I \"" + deploymentType.InstallerName + "\""; }
                    else { installer.InstallCommandLine = deploymentType.InstallString; }
                    installer.ExecutionContext = deploymentType.Context;
                    installer.UserInteractionMode = deploymentType.UserInteractionMode;

                    if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                    {
                        Content content = new Content();
                        ContentFile file = new ContentFile();
                        ContentRef contentRef = new ContentRef();
                        content.Location = deploymentType.ContentPath;
                        file.Name = deploymentType.InstallerName;
                        contentRef.Id = content.Id;
                        content.Files.Add(file);
                        installer.Contents.Add(content);
                        installer.InstallContent = contentRef;

                        if (string.IsNullOrEmpty(deploymentType.UninstallString))
                        {
                            try
                            {
                                string productCode = Utilities.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
                                installer.UninstallCommandLine = "msiexec.exe /X \"" + productCode + "\"";
                            }
                            catch { installer.UninstallCommandLine = null; }
                        }
                    }
                }

                app.DeploymentTypes.Add(dt);
            }

            PublishApplication(siteServer, app);
        }

        public static void CreateAndPublishGeneralApplication(string authoringScope, string applicationName, DeploymentTypes.Enhanced deploymentType, string siteServer, string targetPath)
        {
            ObjectId id = new ObjectId("ScopeId_" + authoringScope, "Application_" + Guid.NewGuid().ToString());
            Application app = new Application(id);
            app.Title = applicationName;
            int appVersion = 1;
            app.Version = appVersion;

            AppDisplayInfo displayInfo = new AppDisplayInfo();
            displayInfo.Title = applicationName;
            displayInfo.Language = CultureInfo.CurrentCulture.Name;
            app.DisplayInfo.Add(displayInfo);

            id = new ObjectId("ScopeId_" + authoringScope, "DeploymentType_" + Guid.NewGuid().ToString());
            if (!string.IsNullOrEmpty(deploymentType.InstallerName) && deploymentType.InstallerName.EndsWith(".msi"))
            {
                DeploymentType dt = new DeploymentType(id, MsiInstaller.TechnologyId);
                dt.Title = "DT_" + applicationName + "_MSI";
                dt.Version = 1;

                MsiInstaller? installer = dt.Installer as MsiInstaller;
                if (installer is not null)
                {
                    GetEnhancedDetectionMethod(app, deploymentType, ref installer);
                    if (string.IsNullOrEmpty(deploymentType.InstallString)) { installer.InstallCommandLine = "msiexec.exe /I \"" + deploymentType.InstallerName + "\""; }
                    else { installer.InstallCommandLine = deploymentType.InstallString; }
                    installer.ExecutionContext = deploymentType.Context;
                    installer.UserInteractionMode = deploymentType.UserInteractionMode;

                    if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                    {
                        Content content = new Content();
                        ContentFile file = new ContentFile();
                        ContentRef contentRef = new ContentRef();
                        content.Location = deploymentType.ContentPath;
                        file.Name = deploymentType.InstallerName;
                        contentRef.Id = content.Id;
                        content.Files.Add(file);
                        installer.Contents.Add(content);
                        installer.InstallContent = contentRef;

                        if (string.IsNullOrEmpty(deploymentType.UninstallString))
                        {
                            try
                            {
                                string productCode = Utilities.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
                                installer.UninstallCommandLine = "msiexec.exe /X \"" + productCode + "\"";
                            }
                            catch { installer.UninstallCommandLine = null; }
                        }
                    }
                }

                app.DeploymentTypes.Add(dt);
            }
            else
            {
                DeploymentType dt = new DeploymentType(id, ScriptInstaller.TechnologyId);
                dt.Title = "DT_" + applicationName + "_Script";
                dt.Version = 1;

                ScriptInstaller? installer = dt.Installer as ScriptInstaller;
                if (installer is not null)
                {
                    GetEnhancedDetectionMethod(app, deploymentType, ref installer);
                    if (string.IsNullOrEmpty(deploymentType.InstallString)) { installer.InstallCommandLine = "msiexec.exe /I \"" + deploymentType.InstallerName + "\""; }
                    else { installer.InstallCommandLine = deploymentType.InstallString; }
                    installer.ExecutionContext = deploymentType.Context;
                    installer.UserInteractionMode = deploymentType.UserInteractionMode;

                    if (!string.IsNullOrEmpty(deploymentType.ContentPath))
                    {
                        Content content = new Content();
                        ContentFile file = new ContentFile();
                        ContentRef contentRef = new ContentRef();
                        content.Location = deploymentType.ContentPath;
                        file.Name = deploymentType.InstallerName;
                        contentRef.Id = content.Id;
                        content.Files.Add(file);
                        installer.Contents.Add(content);
                        installer.InstallContent = contentRef;

                        if (string.IsNullOrEmpty(deploymentType.UninstallString))
                        {
                            try
                            {
                                string productCode = Utilities.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
                                installer.UninstallCommandLine = "msiexec.exe /X \"" + productCode + "\"";
                            }
                            catch { installer.UninstallCommandLine = null; }
                        }
                    }
                }

                app.DeploymentTypes.Add(dt);
            }

            PublishApplication(siteServer, app, targetPath);
        }

        private static void GetEnhancedDetectionMethod(Application app, DeploymentTypes.Enhanced deploymentType, ref MsiInstaller installer)
        {
            installer.DetectionMethod = DetectionMethod.Enhanced;
            switch (deploymentType.DetectionMethod)
            {
                case DetectionMethods.FileOrFolder:
                    DetectionMethods.FileOrFolder? fdm = deploymentType.DetectionMethod as DetectionMethods.FileOrFolder;

                    if (fdm is not null)
                    {
                        object _simpleSetting = new object();
                        if (fdm.FileSystem == DetectionMethods.FileSystemType.File) { _simpleSetting = new FileOrFolder(ConfigurationItemPartType.File, null); }
                        else { _simpleSetting = new FileOrFolder(ConfigurationItemPartType.Folder, null); }
                        FileOrFolder? simpleSetting = _simpleSetting as FileOrFolder;

                        // EnhancedDetectionMethod.Settings[]
                        if (simpleSetting is not null)
                        {
                            simpleSetting.Path = fdm.ObjectPath;
                            simpleSetting.FileOrFolderName = fdm.ObjectName;
                            simpleSetting.Is64Bit = fdm.Is64Bit;
                            simpleSetting.SettingDataType = fdm.ValueType.ValueType;

                            // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                            SettingReference sReference = new SettingReference(
                                app.Scope,
                                app.Name,
                                1,
                                simpleSetting.LogicalName,
                                fdm.ValueType.ValueType,
                                simpleSetting.SourceType,
                                false
                            );
                            sReference.PropertyPath = fdm.Property.ToString();

                            // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                            ConstantValue cValue = new ConstantValue(fdm.Value.ToString(), fdm.ValueType.ValueType);

                            // EnhancedDetectionMethod.Rule.Expressions
                            CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                            cuColl.Add(sReference);
                            cuColl.Add(cValue);
                            Expression expression = new Expression(fdm.Operator.Operator, cuColl);

                            // EnhancedDetectionMethod.Rule
                            Rule rule = new Rule(
                                app.Scope + "/" + app.Name,
                                NoncomplianceSeverity.None,
                                null,
                                expression
                            );

                            // EnhancedDetectionMethod
                            installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                            installer.EnhancedDetectionMethod.Settings.Add(simpleSetting);
                            installer.EnhancedDetectionMethod.Rule = rule;
                        }

                    }
                    break;

                case DetectionMethods.RegistryKey:
                    DetectionMethods.RegistryKey? rdm = deploymentType.DetectionMethod as DetectionMethods.RegistryKey;

                    if (rdm is not null)
                    {
                        if (!rdm.UseDefaultKey & string.IsNullOrEmpty(rdm.ValueName))
                        {
                            RegistryKey? rSimpleSetting = new RegistryKey(null);
                            if (rSimpleSetting is not null)
                            {
                                // EnhancedDetectionMethod.Settings[]
                                rSimpleSetting.RootKey = (RegistryRootKey)rdm.RootKey;
                                rSimpleSetting.Key = rdm.Key;
                                rSimpleSetting.Is64Bit = rdm.Is64Bit;
                                rSimpleSetting.SettingDataType = rdm.ValueType.ValueType;

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                                SettingReference sReference = new SettingReference(
                                    app.Scope,
                                    app.Name,
                                    1,
                                    rSimpleSetting.LogicalName,
                                    rdm.ValueType.ValueType,
                                    rSimpleSetting.SourceType,
                                    false
                                );
                                sReference.PropertyPath = rdm.Property.ToString();

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                                ConstantValue cValue = new ConstantValue(rdm.Value.ToString(), rdm.ValueType.ValueType);

                                // EnhancedDetectionMethod.Rule.Expressions
                                CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                                cuColl.Add(sReference);
                                cuColl.Add(cValue);
                                Expression expression = new Expression(rdm.Operator.Operator, cuColl);

                                // EnhancedDetectionMethod.Rule
                                Rule rule = new Rule(
                                    app.Scope + "/" + app.Name,
                                    NoncomplianceSeverity.None,
                                    null,
                                    expression
                                );

                                // EnhancedDetectionMethod
                                installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                                installer.EnhancedDetectionMethod.Settings.Add(rSimpleSetting);
                                installer.EnhancedDetectionMethod.Rule = rule;
                            }

                        }
                        else
                        {
                            RegistrySetting? rSimpleSetting = new RegistrySetting(null);
                            if (rSimpleSetting is not null)
                            {
                                // EnhancedDetectionMethod.Settings[]
                                rSimpleSetting.RootKey = (RegistryRootKey)rdm.RootKey;
                                rSimpleSetting.Key = rdm.Key;
                                rSimpleSetting.Is64Bit = rdm.Is64Bit;
                                rSimpleSetting.SettingDataType = rdm.ValueType.ValueType;
                                rSimpleSetting.ValueName = rdm.ValueName;

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                                SettingReference sReference = new SettingReference(
                                    app.Scope,
                                    app.Name,
                                    1,
                                    rSimpleSetting.LogicalName,
                                    rdm.ValueType.ValueType,
                                    rSimpleSetting.SourceType,
                                    false
                                );
                                sReference.PropertyPath = rdm.Property.ToString();

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                                ConstantValue cValue = new ConstantValue(rdm.Value.ToString(), rdm.ValueType.ValueType);

                                // EnhancedDetectionMethod.Rule.Expressions
                                CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                                cuColl.Add(sReference);
                                cuColl.Add(cValue);
                                Expression expression = new Expression(rdm.Operator.Operator, cuColl);

                                // EnhancedDetectionMethod.Rule
                                Rule rule = new Rule(
                                    app.Scope + "/" + app.Name,
                                    NoncomplianceSeverity.None,
                                    null,
                                    expression
                                );

                                // EnhancedDetectionMethod
                                installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                                installer.EnhancedDetectionMethod.Settings.Add(rSimpleSetting);
                                installer.EnhancedDetectionMethod.Rule = rule;
                            }
                        }
                    }
                    break;

                case DetectionMethods.MsiEnhanced:
                    DetectionMethods.MsiEnhanced? mdm = deploymentType.DetectionMethod as DetectionMethods.MsiEnhanced;

                    if (mdm is not null)
                    {
                        // EnhancedDetectionMethod.Settings[]
                        MSISettingInstance mSimpleSetting = new MSISettingInstance(mdm.ProductCode, null);
                        mSimpleSetting.SettingDataType = mdm.ValueType.ValueType;

                        // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                        SettingReference sReference = new SettingReference(
                            app.Scope,
                            app.Name,
                            1,
                            mSimpleSetting.LogicalName,
                            mdm.ValueType.ValueType,
                            mSimpleSetting.SourceType,
                            false
                        );
                        sReference.PropertyPath = mdm.Property.ToString();

                        // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                        ConstantValue cValue = new ConstantValue(mdm.Value.ToString(), mdm.ValueType.ValueType);

                        // EnhancedDetectionMethod.Rule.Expressions
                        CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                        cuColl.Add(sReference);
                        cuColl.Add(cValue);
                        Expression expression = new Expression(mdm.Operator.Operator, cuColl);

                        // EnhancedDetectionMethod.Rule
                        Rule rule = new Rule(
                            app.Scope + "/" + app.Name,
                            NoncomplianceSeverity.None,
                            null,
                            expression
                        );

                        // EnhancedDetectionMethod
                        installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                        installer.EnhancedDetectionMethod.Settings.Add(mSimpleSetting);
                        installer.EnhancedDetectionMethod.Rule = rule;
                    }
                    break;
            }
        }

        private static void GetEnhancedDetectionMethod(Application app, DeploymentTypes.Enhanced deploymentType, ref ScriptInstaller installer)
        {
            installer.DetectionMethod = DetectionMethod.Enhanced;
            switch (deploymentType.DetectionMethod)
            {
                case DetectionMethods.FileOrFolder:
                    DetectionMethods.FileOrFolder? fdm = deploymentType.DetectionMethod as DetectionMethods.FileOrFolder;
                    if (fdm is not null)
                    {
                        object _simpleSetting = new object();
                        if (fdm.FileSystem == DetectionMethods.FileSystemType.File) { _simpleSetting = new FileOrFolder(ConfigurationItemPartType.File, null); }
                        else { _simpleSetting = new FileOrFolder(ConfigurationItemPartType.Folder, null); }
                        FileOrFolder? simpleSetting = _simpleSetting as FileOrFolder;

                        // EnhancedDetectionMethod.Settings[]
                        if (simpleSetting is not null)
                        {
                            simpleSetting.Path = fdm.ObjectPath;
                            simpleSetting.FileOrFolderName = fdm.ObjectName;
                            simpleSetting.Is64Bit = fdm.Is64Bit;
                            simpleSetting.SettingDataType = fdm.ValueType.ValueType;

                            // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                            SettingReference sReference = new SettingReference(
                                app.Scope,
                                app.Name,
                                1,
                                simpleSetting.LogicalName,
                                fdm.ValueType.ValueType,
                                simpleSetting.SourceType,
                                false
                            );
                            sReference.PropertyPath = fdm.Property.ToString();

                            // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                            ConstantValue cValue = new ConstantValue(fdm.Value.ToString(), fdm.ValueType.ValueType);

                            // EnhancedDetectionMethod.Rule.Expressions
                            CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                            cuColl.Add(sReference);
                            cuColl.Add(cValue);
                            Expression expression = new Expression(fdm.Operator.Operator, cuColl);

                            // EnhancedDetectionMethod.Rule
                            Rule rule = new Rule(
                                app.Scope + "/" + app.Name,
                                NoncomplianceSeverity.None,
                                null,
                                expression
                            );

                            // EnhancedDetectionMethod
                            installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                            installer.EnhancedDetectionMethod.Settings.Add(simpleSetting);
                            installer.EnhancedDetectionMethod.Rule = rule;
                        }
                    }
                    break;

                case DetectionMethods.RegistryKey:
                    DetectionMethods.RegistryKey? rdm = deploymentType.DetectionMethod as DetectionMethods.RegistryKey;

                    if (rdm is not null)
                    {
                        if (!rdm.UseDefaultKey & string.IsNullOrEmpty(rdm.ValueName))
                        {
                            RegistryKey? rSimpleSetting = new RegistryKey(null);
                            if (rSimpleSetting is not null)
                            {
                                // EnhancedDetectionMethod.Settings[]
                                rSimpleSetting.RootKey = (RegistryRootKey)rdm.RootKey;
                                rSimpleSetting.Key = rdm.Key;
                                rSimpleSetting.Is64Bit = rdm.Is64Bit;
                                rSimpleSetting.SettingDataType = rdm.ValueType.ValueType;

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                                SettingReference sReference = new SettingReference(
                                    app.Scope,
                                    app.Name,
                                    1,
                                    rSimpleSetting.LogicalName,
                                    rdm.ValueType.ValueType,
                                    rSimpleSetting.SourceType,
                                    false
                                );
                                sReference.PropertyPath = rdm.Property.ToString();

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                                ConstantValue cValue = new ConstantValue(rdm.Value.ToString(), rdm.ValueType.ValueType);

                                // EnhancedDetectionMethod.Rule.Expressions
                                CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                                cuColl.Add(sReference);
                                cuColl.Add(cValue);
                                Expression expression = new Expression(rdm.Operator.Operator, cuColl);

                                // EnhancedDetectionMethod.Rule
                                Rule rule = new Rule(
                                    app.Scope + "/" + app.Name,
                                    NoncomplianceSeverity.None,
                                    null,
                                    expression
                                );

                                // EnhancedDetectionMethod
                                installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                                installer.EnhancedDetectionMethod.Settings.Add(rSimpleSetting);
                                installer.EnhancedDetectionMethod.Rule = rule;
                            }

                        }
                        else
                        {
                            RegistrySetting? rSimpleSetting = new RegistrySetting(null);
                            if (rSimpleSetting is not null)
                            {
                                // EnhancedDetectionMethod.Settings[]
                                rSimpleSetting.RootKey = (RegistryRootKey)rdm.RootKey;
                                rSimpleSetting.Key = rdm.Key;
                                rSimpleSetting.Is64Bit = rdm.Is64Bit;
                                rSimpleSetting.SettingDataType = rdm.ValueType.ValueType;
                                rSimpleSetting.ValueName = rdm.ValueName;

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                                SettingReference sReference = new SettingReference(
                                    app.Scope,
                                    app.Name,
                                    1,
                                    rSimpleSetting.LogicalName,
                                    rdm.ValueType.ValueType,
                                    rSimpleSetting.SourceType,
                                    false
                                );
                                sReference.PropertyPath = rdm.Property.ToString();

                                // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                                ConstantValue cValue = new ConstantValue(rdm.Value.ToString(), rdm.ValueType.ValueType);

                                // EnhancedDetectionMethod.Rule.Expressions
                                CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                                cuColl.Add(sReference);
                                cuColl.Add(cValue);
                                Expression expression = new Expression(rdm.Operator.Operator, cuColl);

                                // EnhancedDetectionMethod.Rule
                                Rule rule = new Rule(
                                    app.Scope + "/" + app.Name,
                                    NoncomplianceSeverity.None,
                                    null,
                                    expression
                                );

                                // EnhancedDetectionMethod
                                installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                                installer.EnhancedDetectionMethod.Settings.Add(rSimpleSetting);
                                installer.EnhancedDetectionMethod.Rule = rule;
                            }
                        }
                    }
                    break;

                case DetectionMethods.MsiEnhanced:
                    DetectionMethods.MsiEnhanced? mdm = deploymentType.DetectionMethod as DetectionMethods.MsiEnhanced;

                    if (mdm is not null)
                    {
                        // EnhancedDetectionMethod.Settings[]
                        MSISettingInstance mSimpleSetting = new MSISettingInstance(mdm.ProductCode, null);
                        mSimpleSetting.SettingDataType = mdm.ValueType.ValueType;

                        // EnhancedDetectionMethod.Rule.Expressions.Operands[]<SettingsReference>
                        SettingReference sReference = new SettingReference(
                            app.Scope,
                            app.Name,
                            1,
                            mSimpleSetting.LogicalName,
                            mdm.ValueType.ValueType,
                            mSimpleSetting.SourceType,
                            false
                        );
                        sReference.PropertyPath = mdm.Property.ToString();

                        // EnhancedDetectionMethod.Rule.Expressions.Operands[]<ConstantValue>
                        ConstantValue cValue = new ConstantValue(mdm.Value.ToString(), mdm.ValueType.ValueType);

                        // EnhancedDetectionMethod.Rule.Expressions
                        CustomCollection<ExpressionBase> cuColl = new CustomCollection<ExpressionBase>();
                        cuColl.Add(sReference);
                        cuColl.Add(cValue);
                        Expression expression = new Expression(mdm.Operator.Operator, cuColl);

                        // EnhancedDetectionMethod.Rule
                        Rule rule = new Rule(
                            app.Scope + "/" + app.Name,
                            NoncomplianceSeverity.None,
                            null,
                            expression
                        );

                        // EnhancedDetectionMethod
                        installer.EnhancedDetectionMethod = new Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod();
                        installer.EnhancedDetectionMethod.Settings.Add(mSimpleSetting);
                        installer.EnhancedDetectionMethod.Rule = rule;
                    }
                    break;
            }
        }

        public static void PublishApplication(string siteServer, Application app)
        {
            try
            {
                ManagementScope scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS");
                scope.Connect();
                ObjectQuery? query = new ObjectQuery("Select Name From __NAMESPACE");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                string? siteCode = GetSiteCode(searcher); ;
                scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS\\site_" + siteCode);
                scope.Connect();

                bool appNameCheck = CheckApplicationName(app.Title, scope);
                if (appNameCheck) { throw new ApplicationException("An application with name '" + app.Title + "' already exists."); }

                ManagementClass smsApplication = new ManagementClass(scope, new ManagementPath("SMS_Application"), null);
                ManagementObject application = smsApplication.CreateInstance();

                application.Properties["SDMPackageXML"].Value = SccmSerializer.SerializeToString(app);
                application.Put();

                application.Dispose();
                smsApplication.Dispose();
                searcher.Dispose();

            }
            catch (Exception) { throw; }

        }

        public static void PublishApplication(string siteServer, Application app, string targetPath)
        {
            try
            {
                ManagementScope scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS");
                scope.Connect();
                ObjectQuery? query = new ObjectQuery("Select Name From __NAMESPACE");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                string? siteCode = GetSiteCode(searcher);
                scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS\\site_" + siteCode);
                scope.Connect();

                if (CheckApplicationName(app.Title, scope)) { throw new ApplicationException("An application with name '" + app.Title + "' already exists."); }

                if (!targetPath.StartsWith("Software Library\\Overview\\Application Management\\Applications") & !targetPath.StartsWith("\\Software Library\\Overview\\Application Management\\Applications"))
                {
                    throw new InvalidValueException("Incorrect target path format. Be sure to enter a path using the format: \\Software Library\\Overview\\Application Management\\Applications.");
                }

                targetPath = Regex.Replace(targetPath, "^\\|\\$", "");
                ManagementObject? targetNode = ProcessApplicationPath(targetPath, scope);

                ManagementClass smsApplication = new ManagementClass(scope, new ManagementPath("SMS_Application"), null);
                ManagementObject application = smsApplication.CreateInstance();

                application.Properties["SDMPackageXML"].Value = SccmSerializer.SerializeToString(app);
                application.Put();

                if (targetNode is not null) { MoveApplication(targetNode, application, scope); targetNode.Dispose(); }
                else { throw new ManagementException("Unable to get target folder object. Check if the folder exists and try again."); }

                application.Dispose();
                smsApplication.Dispose();
                searcher.Dispose();

            }
            catch (Exception) { throw; }

        }

        private static string? GetSiteCode(ManagementObjectSearcher searcher)
        {

            foreach (ManagementObject obj in searcher.Get())
            {
                string? smsNamespace = obj.Properties["Name"].Value.ToString();
                if (!string.IsNullOrEmpty(smsNamespace)) { obj.Dispose(); return smsNamespace.Replace("site_", ""); }
            }
            return null;
        }

        private static ManagementObject? ProcessApplicationPath(string path, ManagementScope scope)
        {
            string[] exclusions = new[] { "Software Library", "Overview", "Application Management", "Applications" };
            string[] pathSplit = path.Split("\\");
            string lastNode = pathSplit.Last();
            if (lastNode != "Application")
            {
                foreach (string node in pathSplit)
                {
                    if (!exclusions.Contains(node) && !string.IsNullOrEmpty(node))
                    {
                        ObjectQuery query = new ObjectQuery("Select * From SMS_ObjectContainerNode Where Name = '" + node + "' And ObjectType = 6000");
                        ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                        ManagementObjectCollection result = searcher.Get();
                        if (searcher.Get() is null) { throw new DirectoryNotFoundException("Invalid target path."); }
                        var objNode = from ManagementObject x in result select x;
                        ManagementObject? targetNode = objNode.FirstOrDefault();
                        if (targetNode is not null)
                        {
                            if (targetNode.Properties["Name"].Value.ToString() == lastNode) { return targetNode; }
                        }
                    }
                }
            }
            return null;
        }

        private static void MoveApplication(ManagementObject targetNode, ManagementObject application, ManagementScope scope)
        {
            application.Get();

            ManagementClass smsContainerItem = new ManagementClass(scope, new ManagementPath("SMS_ObjectContainerItem"), null);
            ManagementBaseObject methodParam = smsContainerItem.Methods["MoveMembers"].InParameters;
            methodParam.Properties["InstanceKeys"].Value = new[] { application.Properties["ModelName"].Value };
            methodParam.Properties["ContainerNodeID"].Value = 0;
            methodParam.Properties["ObjectType"].Value = 6000;
            methodParam.Properties["TargetContainerNodeID"].Value = Convert.ToUInt32(targetNode.Properties["ContainerNodeID"].Value);

            smsContainerItem.InvokeMethod("MoveMembers", methodParam, null);
        }

        public static bool CheckApplicationName(string applicationLocalizedName, ManagementScope scope)
        {
            bool output = false;
            ObjectQuery query = new ObjectQuery("Select ModelName From SMS_ApplicationLatest Where LocalizedDisplayName = '" + applicationLocalizedName + "'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection result = searcher.Get();
            if (result.Count > 0)
            {
                searcher.Dispose();
                result.Dispose();
                output = true;
                return output;
            }
            searcher.Dispose();
            return output;
        }

    }
}