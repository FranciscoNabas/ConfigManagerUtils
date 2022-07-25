# Configuration Manager Utilities

Or ConfigManagerUtils for short, is a class library designed to facilitate automation and management for Microsoft Endpoint Configuration Manager.
This tool is being built primarily to be used on PowerShell or other shells that accept dotnet integration.

This tool is still under expansion, what do you like to see on it?

*Disclaimer:* I'm not a developer, I'm just a curious Systems Administrator who LOVES PowerShell and ConfigMgr.
This code is far from being the best way of doing it.
Suggestions are welcome!

## Import methods

```powershell
Add-Type -Path '.\ConfigManagerUtils.dll'
# Or
[void][System.Reflection.Assembly]::LoadFrom('.\ConfigManagerUtils.dll')
```

## List of contents

### Utilities

-   Search Console Objects

```powershell
PS>_ $result = [ConfigManagerUtils.Utilities.Console]::SearchObject('Application Name', [ObjectContainerType]::Application, 'SITESERVER.contoso.com')
PS>_ $result

@"
Name        : Application Name
InstanceKey : AppModelName
Type        : Application
Path        : \Application\Path\Example
"@

PS>_ 
```

-   Get MSI Product Code

```powershell
[ConfigManagerUtils.Utilities.Software]::GetMsiProductCode('.\ProductInstaller.msi')

'{F599B4ED-316F-4FFF-9AA9-78A1DAE31CEA}'
```

-   Get formatted Last Win32 Error

```powershell
$errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
[ConfigManagerUtils.Utilities.Software]::GetFormatedWin32Error($errorCode)

'The system could not find the environment option that was entered.'
```

### Status Messages

-   Format Status Messages

```powershell
[StatusMessages.StatusMessage]::GetFormatedMessage($msgId, [Severity]::Error, [Module]::SMSServer, $insertStrings)

$statusObject = [StatusMessages.StatusMessage]::new($msgId, [Severity]::Error, [Module]::SMSServer, $insertStrings)
[StatusMessages.StatusMessage]::GetFormatedMessage($statusObject)
```

### Application Management

-   Detection Method creation.

```powershell
# Simple file detection clause example.
$contentPaht = 'C:\Program Files\YourInstalledSoftware'
$fileName = 'Software.exe'
$fileSystemType = [ConfigManagerUtils.Applications.DetectionMethods.FileSystemType]::File

$detectionMethod = New-Object ConfigManagerUtils.Applications.DetectionMethods.FileOrFolder -ArgumentList $contentPath, $fileName, $fileSystemType
# Or
$detectionMethod = [ConfigManagerUtils.Applications.DetectionMethods.FileOrFolder]::new($contentPath, $fileName, $fileSystemType)

# File attribute comparison clause example.
$contentPaht = 'C:\Program Files\YourInstalledSoftware'
$fileName = 'Software.exe'
$property = [ConfigManagerUtils.ObjectProperty]::DateCreated
$fileSystemType = [ConfigManagerUtils.Applications.DetectionMethods.FileSystemType]::File
$operator = [ConfigManagerUtils.ExpressionOperator]::GreaterThan
$value = '7/25/2022'

$detectionMethod = New-Object ConfigManagerUtils.Applications.DetectionMethods.FileOrFolder -ArgumentList $contentPath, $fileName, $property, $operator, $value, $fileSystemType
# Or
$detectionMethod = [ConfigManagerUtils.Applications.DetectionMethods.FileOrFolder]::new($contentPath, $fileName, $property, $operator, $value, $fileSystemType)
```

-   Deployment Type Creation

```powershell
$dm = [FileOrFolder]::new('C:\Program Files\YourInstalledSoftware','Software.exe', [FileSystemType]::File)
$depType = [DeploymentTypes.Enhanced]::new('\\APPREPO.contoso.com\Software\YourSoftware','YourInstaller.msi', 'msiexec.exe /I YourInstaller.msi')
```

- Simple MSI Application Creation

```powershell
$depType = [DeploymentTypes.Msi]::new('\\APPREPO.contoso.com\Software\YourSoftware','YourInstaller.msi', '/qn /norestart /l*vx "C:\logs.log"')
[Applications.Application]::CreateAndPublishMsiApplication($siteAuthoringScopeCode, 'Application Name', $depType, 'SITESERVER.contoso.com')
```