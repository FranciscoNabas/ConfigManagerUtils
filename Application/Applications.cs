using System.Globalization;
using System.Text.RegularExpressions;
using System.Management;
using Microsoft.ConfigurationManagement.ApplicationManagement;
using Microsoft.ConfigurationManagement.ApplicationManagement.Serialization;
using Microsoft.ConfigurationManagement.DesiredConfigurationManagement;
using Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions;
using Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules;
using ConfigManagerUtils.Applications.DeploymentTypes;
using ConfigManagerUtils.Utilities;

namespace ConfigManagerUtils.Applications
{
    

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
                                string productCode = Software.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
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
                                string productCode = Software.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
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
                                string productCode = Software.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
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
                                string productCode = Software.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
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
                                string productCode = Software.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
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
                                string productCode = Software.GetMsiProductCode(deploymentType.ContentPath + deploymentType.InstallerName);
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

                string? siteCode = ConfigManagerUtils.Utilities.Console.GetSiteCode(searcher);
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

                string? siteCode = ConfigManagerUtils.Utilities.Console.GetSiteCode(searcher);
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
                        if (result is null) { throw new DirectoryNotFoundException("Invalid target path."); }
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