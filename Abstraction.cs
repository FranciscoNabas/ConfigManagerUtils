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
    internal abstract class DataTypeEnumeration : IComparable
    {
        
        internal int Id { get; private set; }
        internal string Name { get; private set; }
        internal Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType ValueType { get; private set; }

        protected DataTypeEnumeration(int id, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType valueType) => (Id, ValueType, Name) = (id, valueType, valueType.Name);

        internal static IEnumerable<T> GetAll<T>() where T : DataTypeEnumeration =>
            typeof(T).GetFields(BindingFlags.Public |
                                BindingFlags.Static |
                                BindingFlags.DeclaredOnly)
                     .Select(f => f.GetValue(null))
                     .Cast<T>();

        public override bool Equals(object? obj)
        {
            if (obj is not DataTypeEnumeration otherValue)
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
            else { return Id.CompareTo(((DataTypeEnumeration)other).Id); }
        }
        
    }
    internal abstract class ExpressionOperatorEnumeration : IComparable
    {
        internal int Id { get; private set; }
        internal string Name { get; private set; }
        internal Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator Operator { get; private set; }

        protected ExpressionOperatorEnumeration(int id, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator _operator) => (Id, Operator, Name) = (id, _operator, _operator.OperatorName);

        public override string ToString() => Name;

        internal static IEnumerable<T> GetAll<T>() where T : ExpressionOperatorEnumeration =>
            typeof(T).GetFields(BindingFlags.Public |
                                BindingFlags.Static |
                                BindingFlags.DeclaredOnly)
                     .Select(f => f.GetValue(null))
                     .Cast<T>();

        public override bool Equals(object? obj)
        {
            if (obj is not ExpressionOperatorEnumeration otherValue)
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
            else { return Id.CompareTo(((ExpressionOperatorEnumeration)other).Id); }
        }
    }
    internal abstract class Enumeration : IComparable
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
    internal class InternModule : Enumeration
    {
        public static InternModule SMSClient => new(1, "climsgs.dll");
        public static InternModule SMSProvider => new(2, "provmsgs.dll");
        public static InternModule SMSServer => new(3, "srvmsgs.dll");
        public InternModule(int Id, string Name) : base(Id, Name) { }
    }
    internal class InternDataType : DataTypeEnumeration
    {
        public static InternDataType Boolean => new(1, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Boolean);
        public static InternDataType DateTime => new(2, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.DateTime);
        public static InternDataType CIMDateTime => new(3, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.CIMDateTime);
        public static InternDataType Double => new(4, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Double);
        public static InternDataType Int64 => new(5, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Int64);
        public static InternDataType String => new(6, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.String);
        public static InternDataType Version => new(7, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Version);
        public static InternDataType FileSystemAccessControl => new(8, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAccessControl);
        public static InternDataType RegistryAccessControl => new(9, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.RegistryAccessControl);
        public static InternDataType FileSystemAttribute => new(10, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAttribute);
        public static InternDataType Base64 => new(11, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Base64);
        public static InternDataType XML => new(12, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.XML);
        public static InternDataType Other => new(13, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Other);
        public static InternDataType Complex => new(14, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Complex);
        public static InternDataType BooleanArray => new(15, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.BooleanArray);
        public static InternDataType DateTimeArray => new(16, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.DateTimeArray);
        public static InternDataType CIMDateTimeArray => new(17, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.CIMDateTimeArray);
        public static InternDataType DoubleArray => new(18, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.DoubleArray);
        public static InternDataType Int64Array => new(19, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.Int64Array);
        public static InternDataType StringArray => new(20, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.StringArray);
        public static InternDataType VersionArray => new(21, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.VersionArray);
        public static InternDataType FileSystemAccessControlArray => new(22, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAccessControlArray);
        public static InternDataType RegistryAccessControlArray => new(23, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.RegistryAccessControlArray);
        public static InternDataType FileSystemAttributeArray => new(24, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType.FileSystemAttributeArray);

        public InternDataType(int id, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType valueType) : base(id, valueType) { }
    }
    internal class InternExpressionOperator : ExpressionOperatorEnumeration
    {
        public static InternExpressionOperator And => new(1, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.And);
        public static InternExpressionOperator Or => new(2, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.Or);
        public static InternExpressionOperator IsEquals => new(3, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.IsEquals);
        public static InternExpressionOperator NotEquals => new(4, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotEquals);
        public static InternExpressionOperator GreaterThan => new(5, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.GreaterThan);
        public static InternExpressionOperator LessThan => new(6, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.LessThan);
        public static InternExpressionOperator Between => new(7, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.Between);
        public static InternExpressionOperator GreaterEquals => new(8, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.GreaterEquals);
        public static InternExpressionOperator LessEquals => new(9, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.LessEquals);
        public static InternExpressionOperator BeginsWith => new(10, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.BeginsWith);
        public static InternExpressionOperator NotBeginsWith => new(11, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotBeginsWith);
        public static InternExpressionOperator EndsWith => new(12, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.EndsWith);
        public static InternExpressionOperator NotEndsWith => new(13, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotEndsWith);
        public static InternExpressionOperator Contains => new(14, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.Contains);
        public static InternExpressionOperator NotContains => new(15, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NotContains);
        public static InternExpressionOperator AllOf => new(16, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.AllOf);
        public static InternExpressionOperator OneOf => new(17, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.OneOf);
        public static InternExpressionOperator NoneOf => new(18, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.NoneOf);
        public static InternExpressionOperator SetEquals => new(19, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.SetEquals);
        public static InternExpressionOperator SubsetOf => new(20, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.SubsetOf);
        public static InternExpressionOperator ExcludesAll => new(21, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator.ExcludesAll);

        public InternExpressionOperator(int id, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator _operator)
            : base(id, _operator) { }
    }
    public class DataType
    {
        private int Id { get; }
        public Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType ValueType { get; }

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

        private DataType(int id, Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType valueType)
        {
            InternDataType idt = new InternDataType(id, valueType);
            Id = idt.Id;
            ValueType = idt.ValueType;
        }
    }
    
    public class ExpressionOperator
    {
        private int Id { get;}
        public Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator Operator { get; }
        
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

        private ExpressionOperator(int id, Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator _operator)
        {
            InternExpressionOperator ieo = new InternExpressionOperator(id, _operator);
            Id = ieo.Id;
            Operator = ieo.Operator;
        }
    }
    public abstract class EnhancedDetectionMethod
    {
        internal bool Is64Bit { get; set; }

        internal EnhancedDetectionMethod() { Is64Bit = true; }
        
    }

}