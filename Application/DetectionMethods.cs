using Microsoft.ConfigurationManagement.DesiredConfigurationManagement;
using ConfigManagerUtils.Utilities;
using System;
using System.IO;

#nullable enable
namespace ConfigManagerUtils.Applications.DetectionMethods
{
    public enum RegistryRootKey
    {
        ClassesRoot = Microsoft.ConfigurationManagement.DesiredConfigurationManagement.RegistryRootKey.ClassesRoot,
        CurrentConfig = Microsoft.ConfigurationManagement.DesiredConfigurationManagement.RegistryRootKey.CurrentConfig,
        CurrentUser = Microsoft.ConfigurationManagement.DesiredConfigurationManagement.RegistryRootKey.CurrentUser,
        LocalMachine = Microsoft.ConfigurationManagement.DesiredConfigurationManagement.RegistryRootKey.LocalMachine,
        Users = Microsoft.ConfigurationManagement.DesiredConfigurationManagement.RegistryRootKey.Users
    }

    public enum FileSystemType
    {
        File,
        Folder
    }

    public class FileOrFolder : EnhancedDetectionMethod
    {
        public string? Location
        {
            get
            {
                if (!string.IsNullOrEmpty(ObjectPath))
                {
                    if (ObjectPath.EndsWith("\\", StringComparison.Ordinal))
                    {
                        return ObjectPath + ObjectName;
                    }
                    return ObjectPath + "\\" + ObjectName;
                }
                else { return null; }

            }
        }
        public string ObjectPath { get; set; }
        public string? ObjectName { get; set; }
        public object Value { get; set; }
        public FileSystemType FileSystem { get; set; }
        public ObjectProperty? Property { get; set; }
        public ExpressionOperator Operator { get; set; }
        public DataType ValueType { get; set; }

        public FileOrFolder(string objectPath, string objectName, FileSystemType fileSystem)
        {
            if (!objectPath.EndsWith("\\", StringComparison.Ordinal)) { objectPath = objectPath + "\\"; }
            Software.ValidadeLocalPath(objectPath + objectName);

            FileSystem = fileSystem;
            ObjectPath = objectPath;
            ObjectName = objectName;
            Operator = ExpressionOperator.NotEquals;
            ValueType = DataType.Int64;
            Value = new Int64();
        }
        public FileOrFolder(string objectPath, string objectName, ObjectProperty property, ExpressionOperator _operator, string value, FileSystemType fileSystem)
        {
            if (!objectPath.EndsWith("\\", StringComparison.Ordinal)) { objectPath = objectPath + "\\"; }
            Software.ValidadeLocalPath(objectPath + objectName);

            if (property == ObjectProperty.RegistryKeyExists || property == ObjectProperty.ProductVersion) {
                throw new InvalidDataTypeException("Invalid data type. Cannot use RegistryKeyExists or ProductVersion with File or Folder.");
            }

            object _value = new object();
            DataType _valueType = DataType.Int64;
            switch (property)
            {
                case ObjectProperty.DateModified | ObjectProperty.DateCreated:
                    try {
                        DateTime dateTime = DateTime.Parse(value).ToUniversalTime();
                        _value = string.Format("{0:O}", dateTime, dateTime.Kind);
                        _valueType = DataType.DateTime;
                    }
                    catch (Exception) { throw; }
                    break;

                case ObjectProperty.Version:
                    try { _value = Version.Parse(value); _valueType = DataType.Version; } catch { throw new FormatException("Wrong format. Make sure the value is a valid Version."); }
                    break;

                case ObjectProperty.Size:
                    try { _value = Int64.Parse(value); _valueType = DataType.Int64; } catch { throw new FormatException("Wrong format. Make sure the value is a valid Integer."); }
                    break;
            }
            FileSystem = fileSystem;
            Value = _value;
            ValueType = _valueType;
            Property = property;
            Operator = _operator;
            ObjectPath = objectPath;
            ObjectName = objectName;
        }
    }
    public class RegistryKey : EnhancedDetectionMethod
    {

        public string Location => DCMUtils.GetHiveString((Microsoft.ConfigurationManagement.DesiredConfigurationManagement.RegistryRootKey)RootKey) + "\\" + Key;
        public RegistryRootKey RootKey { get; set; }
        public string Key { get; set; }
        public string? ValueName { get; set; }
        public object Value { get; set; }
        public bool UseDefaultKey { get; set; }
        public ObjectProperty? Property { get; set; }
        public ExpressionOperator Operator { get; set; }
        public DataType ValueType { get; set; }

        public RegistryKey(RegistryRootKey rootKey, string key)
        {
            RootKey = rootKey;
            Key = key;
            Property = ObjectProperty.RegistryKeyExists;
            Operator = ExpressionOperator.IsEquals;
            ValueType = DataType.Boolean;
            Value = true;
            UseDefaultKey = false;

        }
        public RegistryKey(RegistryRootKey rootKey, string key, DataType valueType, bool useDefaultKey = true)
        {
            RootKey = rootKey;
            Key = key;
            Property = ObjectProperty.RegistryKeyExists;
            Operator = ExpressionOperator.IsEquals;
            ValueType = valueType;
            UseDefaultKey = true;

            object _value = new object();
            if (valueType.ValueType.TypeName != "Version" & valueType.ValueType.TypeName != "Int64" & valueType.ValueType.TypeName != "String") {
                throw new InvalidDataTypeException("Invalid data type. You can only use Version, Int64 or String with RegistryKey.");
            }
            switch (valueType.ValueType.TypeName)
            {
                case "Version":
                    try { _value = Version.Parse("0.0"); } catch { throw new FormatException("Wrong format. Make sure the value is a valid Version."); }
                    break;

                case "Int64":
                    try { _value = Int64.Parse("0"); } catch { throw new FormatException("Wrong format. Make sure the value is a valid Integer."); }
                    break;

                case "String":
                    _value = new string('N', 'A');
                    break;
            }
            Value = _value;

        }
        public RegistryKey(RegistryRootKey rootKey, string key, DataType valueType, ExpressionOperator _operator, string valueName, string value)
        {
            object _value = new object();
            if (valueType.ValueType.TypeName != "Version" & valueType.ValueType.TypeName != "Int64" & valueType.ValueType.TypeName != "String") {
                throw new InvalidDataTypeException("Invalid data type. You can only use Version, Int64 or String with RegistryKey.");
            }
            switch (valueType.ValueType.TypeName)
            {
                case "Version":
                    try { _value = Version.Parse(value); } catch { throw new FormatException("Wrong format. Make sure the value is a valid Version."); }
                    break;

                case "Int64":
                    try { _value = Int64.Parse(value); } catch { throw new FormatException("Wrong format. Make sure the value is a valid Integer."); }
                    break;

                case "String":
                    _value = value;
                    break;
            }
            Operator = _operator;
            ValueName = valueName;
            Value = _value;
            ValueType = valueType;
            RootKey = rootKey;
            Key = key;
            UseDefaultKey = false;
        }
        public RegistryKey(RegistryRootKey rootKey, string key, DataType valueType, ExpressionOperator _operator, string? valueName, string value, bool useDefaultKey = true)
        {
            object _value = new object();
            if (valueType.ValueType.TypeName != "Version" & valueType.ValueType.TypeName != "Int64" & valueType.ValueType.TypeName != "String") {
                throw new InvalidDataTypeException("Invalid data type. You can only use Vertion, Int64 or String with RegistryKey.");
            }
            switch (valueType.ValueType.TypeName)
            {
                case "Version":
                    try { _value = Version.Parse(value); } catch { throw new FormatException("Wrong format. Make sure the value is a valid Version."); }
                    break;

                case "Int64":
                    try { _value = Int64.Parse(value); } catch { throw new FormatException("Wrong format. Make sure the value is a valid Integer."); }
                    break;

                case "String":
                    _value = value;
                break;
            }
            Operator = _operator;
            ValueName = valueName;
            Value = _value;
            ValueType = valueType;
            RootKey = rootKey;
            Key = key;
            UseDefaultKey = true;
        }
    }
    public class MsiEnhanced : EnhancedDetectionMethod
    {
        public string FilePath { get; set; }
        public string ProductCode { get; set; }
        public object Value { get; set; }
        public ObjectProperty? Property { get; set; }
        public ExpressionOperator Operator { get; set; }
        public DataType ValueType { get; set; }

        public MsiEnhanced(string filePath)
        {
            if (!File.Exists(filePath)) { throw new FileNotFoundException("File not found. Make sure the file path entered exists and it's accessible."); }
            string productCode = Software.GetMsiProductCode(filePath);

            FilePath = filePath;
            ProductCode = productCode;
            Operator = ExpressionOperator.NotEquals;
            ValueType = DataType.Int64;
            Value = new Int64();

        }
        public MsiEnhanced(string filePath, ExpressionOperator _operator, Version value)
        {
            if (!File.Exists(filePath)) { throw new FileNotFoundException("File not found. Make sure the file path entered exists and it's accessible."); }
            string productCode = Software.GetMsiProductCode(filePath);

            FilePath = filePath;
            ProductCode = productCode;
            Operator = _operator;
            Value = value;
            ValueType = DataType.Version;
            Property = ObjectProperty.ProductVersion;

        }
    }

}