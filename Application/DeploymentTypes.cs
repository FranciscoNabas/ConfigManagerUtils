using System.Management.Automation;
using Microsoft.ConfigurationManagement.ApplicationManagement;
using ConfigManagerUtils.Utilities;
using System.IO;
using System;

#nullable enable
namespace ConfigManagerUtils.Applications.DeploymentTypes
{
    public enum ScriptLanguage
    {
        PowerShell,
        VBScript,
        JavaScript
    }
    public class Msi
    {
        public string ContentPath { get; set; }
        public string InstallerName { get; set; }
        public string? MsiInstallArgs { get; set; }
        public string? InstallString { get; }
        public string? UninstallString { get; set; }
        internal string ProductCode { get; }
        public Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext Context { get; set; }
        public UserInteractionMode UserInteractionMode { get; set; }
        
        public Msi(string contentPath, string installerName, string msiInstallArgs)
        {
            if (string.IsNullOrEmpty(msiInstallArgs)) { InstallString = "msiexec.exe /I " + installerName; }
            else { InstallString = "msiexec.exe /I " + installerName + " " + msiInstallArgs; }

            if (!contentPath.EndsWith("\\", StringComparison.Ordinal)) { contentPath = contentPath + "\\"; }
            if (!File.Exists(contentPath + installerName)) { throw new FileNotFoundException("File not found. Make sure the file path entered exists and it's accessible."); }

            string productCode = Software.GetMsiProductCode(contentPath + installerName);
            UninstallString = "msiexec.exe /X " + productCode.Trim() + " /qn /norestart";

            ProductCode = productCode;
            ContentPath = contentPath;
            InstallerName = installerName;
            MsiInstallArgs = msiInstallArgs;
            Context = Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext.System;
            UserInteractionMode = UserInteractionMode.Hidden;
        }
        public Msi(string contentPath, string installerName, string msiInstallArgs, Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext executionContext, UserInteractionMode userInteractionMode)
        {
            if (!string.IsNullOrEmpty(msiInstallArgs)) { InstallString = "msiexec.exe /I " + installerName; }
            else { InstallString = "msiexec.exe /I " + installerName + " " + msiInstallArgs; }

            if (contentPath.EndsWith("\\", StringComparison.Ordinal) == false) { contentPath = contentPath + "\\"; }
            if (!File.Exists(contentPath + installerName)) { throw new FileNotFoundException("File not found. Make sure the file path entered exists and it's accessible."); }

            string productCode = Software.GetMsiProductCode(contentPath + installerName);
            UninstallString = "msiexec.exe /X " + productCode.Trim() + " /qn /norestart";

            ProductCode = productCode;
            ContentPath = contentPath;
            InstallerName = installerName;
            MsiInstallArgs = msiInstallArgs;
            Context = executionContext;
            UserInteractionMode = userInteractionMode;
        }
    }
    public class Script
    {
        public string? ContentPath { get; set; }
        public string? InstallerName { get; set; }
        public string InstallString { get; set; }
        public string? UninstallString { get; set; }
        public Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext Context { get; set; }
        public UserInteractionMode UserInteractionMode { get; set; }
        public bool RunAs32Bit { get; set; }
        public ScriptLanguage Language { get; set; }
        public ScriptBlock ScriptBlock { get; set; }

        public Script(ScriptLanguage language, ScriptBlock scriptBlock, string installString)
        {
            InstallString = installString;
            Language = language;
            ScriptBlock = scriptBlock;
            RunAs32Bit = false;
            Context = Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext.System;
            UserInteractionMode = UserInteractionMode.Hidden;
        }
        public Script(string contentPath, string installerName, string installString, ScriptLanguage language, ScriptBlock scriptBlock)
        {
            if (!contentPath.EndsWith("\\", StringComparison.Ordinal)) { contentPath = contentPath + "\\"; }
            if (!File.Exists(contentPath + installerName)) { throw new FileNotFoundException("File not found. Make sure the file path entered exists and it's accessible."); }
            
            ContentPath = contentPath;
            InstallerName = installerName;
            InstallString = installString;
            Language = language;
            ScriptBlock = scriptBlock;
            RunAs32Bit = false;
            Context = Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext.System;
            UserInteractionMode = UserInteractionMode.Hidden;
        }
    }
    public class Enhanced
    {
        public string? ContentPath { get; set; }
        public string? InstallerName { get; set; }
        public string InstallString { get; set; }
        public string? UninstallString { get; set; }
        public Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext Context { get; set; }
        public UserInteractionMode UserInteractionMode { get; set; }
        public EnhancedDetectionMethod DetectionMethod { get; set; }

        public Enhanced(string installString, EnhancedDetectionMethod detectionMethod)
        {
            InstallString = installString;
            DetectionMethod = detectionMethod;
            Context = Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext.System;
            UserInteractionMode = UserInteractionMode.Hidden;
        }
        public Enhanced(string contentPath, string installerName, string installString, EnhancedDetectionMethod detectionMethod)
        {
            if (!contentPath.EndsWith("\\", StringComparison.Ordinal)) { contentPath = contentPath + "\\"; }
            if (!File.Exists(contentPath + installerName)) { throw new FileNotFoundException("File not found. Make sure the file path entered exists and it's accessible."); }
            
            ContentPath = contentPath;
            InstallerName = installerName;
            InstallString = installString;
            DetectionMethod = detectionMethod;
            Context = Microsoft.ConfigurationManagement.ApplicationManagement.ExecutionContext.System;
            UserInteractionMode = UserInteractionMode.Hidden;
        }
    }
}