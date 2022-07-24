using System.Reflection;
using System.Management.Automation;
using System.ComponentModel;
using Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators;
using Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions;
using ConfigManagerUtils.Applications;

namespace ConfigManagerUtils
{
    public enum ObjectProperty
    {
        DateModified,
        DateCreated,
        Version,
        Size,
        RegistryKeyExists,
        ProductVersion
    }
    
    public abstract class Enumeration : IComparable
    {
        internal string Name { get; set; }

        internal int Id { get; set; }

        protected Enumeration(int id, string name) => (Id, Name) = (id, name);

        public override string ToString() => Name;

        internal static IEnumerable<T> GetAll<T>() where T : Enumeration =>
            typeof(T).GetFields(BindingFlags.Public |
                                BindingFlags.Static |
                                BindingFlags.DeclaredOnly)
                     .Select(f => f.GetValue(null))
                     .Cast<T>();

        public override bool Equals(object? obj)
        {
            if (obj is not Enumeration otherValue)
            {
                return false;
            }

            var typeMatches = GetType().Equals(obj.GetType());
            var valueMatches = Id.Equals(otherValue.Id);

            return typeMatches && valueMatches;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Name, Id);
        }

        public int CompareTo(object? other)
        {
            if (other == null) { return 0; }
            else { return Id.CompareTo(((Enumeration)other).Id); }
        }

        // Other utility methods ...
    }
    
    public class Module : Enumeration
    {
        public static Module SMSClient => new(1, "climsgs.dll");
        public static Module SMSProvider => new(2, "provmsgs.dll");
        public static Module SMSServer => new(3, "srvmsgs.dll");
        public Module(int Id, string Name) : base(Id, Name) { }
    }
    
    public abstract class EnhancedDetectionMethod
    {
        internal bool Is64Bit { get; set; }

        internal EnhancedDetectionMethod() { Is64Bit = true; }
        
    }

    #region ObjectContainerType
    public abstract class ObjectContainerEnumeration : Enumeration
    {
        internal string UnderlyingClassName { get; private set; }

        internal ObjectContainerEnumeration(int id, string name, string underlyingClassName)
            : base(id, name)
        { UnderlyingClassName = underlyingClassName; }
    }

    public class ObjectContainerType : ObjectContainerEnumeration
    {
        public static ObjectContainerType Package => new(2, "Package", "SMS_Package");
        public static ObjectContainerType Advertisement => new(3, "Advertisement", "SMS_Advertisement");
        public static ObjectContainerType Query => new(7, "Query", "SMS_Query");
        public static ObjectContainerType Report => new(8, "Report", "SMS_Report");
        public static ObjectContainerType MeteredProductRule => new(9, "MeteredProductRule", "SMS_MeteredProductRule");
        public static ObjectContainerType ConfigurationItem => new(11, "ConfigurationItem", "SMS_ConfigurationItem");
        public static ObjectContainerType OperatingSystemInstallPackage => new(14, "OperatingSystemInstallPackage", "SMS_OperatingSystemInstallPackage");
        public static ObjectContainerType StateMigration => new(17, "StateMigration", "SMS_StateMigration");
        public static ObjectContainerType ImagePackage => new(18, "ImagePackage", "SMS_ImagePackage");
        public static ObjectContainerType BootImagePackage => new(19, "BootImagePackage", "SMS_BootImagePackage");
        public static ObjectContainerType TaskSequence => new(20, "TaskSequence", "SMS_TaskSequencePackage");
        public static ObjectContainerType DeviceSettingPackage => new(21, "DeviceSettingPackage", "SMS_DeviceSettingPackage");
        public static ObjectContainerType DriverPackage => new(23, "DriverPackage", "SMS_DriverPackage");
        public static ObjectContainerType Driver => new(25, "Driver", "SMS_Driver");
        public static ObjectContainerType SoftwareUpdate => new(1011, "SoftwareUpdate", "SMS_SoftwareUpdate");
        public static ObjectContainerType ConfigurationBaseline => new(2011, "ConfigurationBaseline", "SMS_ConfigurationItem");
        public static ObjectContainerType DeviceCollection => new(5000, "DeviceCollection", "SMS_Collection");
        public static ObjectContainerType UserCollection => new(5001, "UserCollection", "SMS_Collection");
        public static ObjectContainerType Application => new(6000, "Application", "SMS_ApplicationLatest");
        public static ObjectContainerType ConfigurationItem_ComplianceSettings => new(6001, "ConfigurationItem_ComplianceSettings", "SMS_ConfigurationItemLatest");

        public ObjectContainerType(int id, string name, string underlyingClassName) : base(id, name, underlyingClassName) {}
    }
    #endregion
    
    #region  DataType
    public abstract class DataTypeEnumeration : Enumeration
    {
        internal Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType ValueType { get; private set; }

        protected DataTypeEnumeration(int id, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType valueType)
            : base(id, valueType.Name)
        { ValueType = valueType; }
        
    }
    
    public class DataType : DataTypeEnumeration
    {
        public static DataType Boolean => new(1, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Boolean);
        public static DataType DateTime => new(2, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.DateTime);
        public static DataType CIMDateTime => new(3, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.CIMDateTime);
        public static DataType Double => new(4, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Double);
        public static DataType Int64 => new(5, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Int64);
        public static DataType String => new(6, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.String);
        public static DataType Version => new(7, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Version);
        public static DataType FileSystemAccessControl => new(8, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAccessControl);
        public static DataType RegistryAccessControl => new(9, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.RegistryAccessControl);
        public static DataType FileSystemAttribute => new(10, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAttribute);
        public static DataType Base64 => new(11, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Base64);
        public static DataType XML => new(12, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.XML);
        public static DataType Other => new(13, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Other);
        public static DataType Complex => new(14, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Complex);
        public static DataType BooleanArray => new(15, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.BooleanArray);
        public static DataType DateTimeArray => new(16, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.DateTimeArray);
        public static DataType CIMDateTimeArray => new(17, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.CIMDateTimeArray);
        public static DataType DoubleArray => new(18, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.DoubleArray);
        public static DataType Int64Array => new(19, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Int64Array);
        public static DataType StringArray => new(20, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.StringArray);
        public static DataType VersionArray => new(21, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.VersionArray);
        public static DataType FileSystemAccessControlArray => new(22, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAccessControlArray);
        public static DataType RegistryAccessControlArray => new(23, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.RegistryAccessControlArray);
        public static DataType FileSystemAttributeArray => new(24, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAttributeArray);

        public DataType(int id, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType valueType) : base(id, valueType) { }
    }
    #endregion
    
    #region ExpressionOperator
    public abstract class ExpressionOperatorEnumeration : Enumeration
    {
        internal Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator Operator { get; private set; }

        protected ExpressionOperatorEnumeration(int id, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator _operator)
            : base(id, _operator.OperatorName)
        { Operator = _operator; }

    }

    public class ExpressionOperator : ExpressionOperatorEnumeration
    {
        public static ExpressionOperator And => new(1, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.And);
        public static ExpressionOperator Or => new(2, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.Or);
        public static ExpressionOperator IsEquals => new(3, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.IsEquals);
        public static ExpressionOperator NotEquals => new(4, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotEquals);
        public static ExpressionOperator GreaterThan => new(5, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.GreaterThan);
        public static ExpressionOperator LessThan => new(6, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.LessThan);
        public static ExpressionOperator Between => new(7, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.Between);
        public static ExpressionOperator GreaterEquals => new(8, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.GreaterEquals);
        public static ExpressionOperator LessEquals => new(9, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.LessEquals);
        public static ExpressionOperator BeginsWith => new(10, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.BeginsWith);
        public static ExpressionOperator NotBeginsWith => new(11, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotBeginsWith);
        public static ExpressionOperator EndsWith => new(12, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.EndsWith);
        public static ExpressionOperator NotEndsWith => new(13, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotEndsWith);
        public static ExpressionOperator Contains => new(14, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.Contains);
        public static ExpressionOperator NotContains => new(15, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotContains);
        public static ExpressionOperator AllOf => new(16, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.AllOf);
        public static ExpressionOperator OneOf => new(17, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.OneOf);
        public static ExpressionOperator NoneOf => new(18, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NoneOf);
        public static ExpressionOperator SetEquals => new(19, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.SetEquals);
        public static ExpressionOperator SubsetOf => new(20, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.SubsetOf);
        public static ExpressionOperator ExcludesAll => new(21, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.ExcludesAll);

        public ExpressionOperator(int id, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator _operator)
            : base(id, _operator) { }
    }
    #endregion

}