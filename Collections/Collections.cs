using System.Management;
using ConfigManagerUtils;

namespace ConfigManagerUtils.Collections
{
    public enum RuleType
    {
        Direct,
        Exclude,
        Include,
        Query
    }
    public enum ResourceType
    {
        Device,
        User
    }
    public enum CollectionStatus : uint
    {
        None = 0,
        Ready = 1,
        Refreshing = 2,
        Saving = 3,
        Evaluating = 4,
        OnRefreshQueue = 5,
        Deleting = 6,
        AddingMember = 7,
        Querying = 8
    }
    public class DirectRule : CollectionRule
    {
        public ResourceType? ResourceType { get; set; }
        public uint[]? ResourceId { get; set; }
        public DirectRule(RuleType type = RuleType.Direct) : base(type) { }
        public DirectRule(ResourceType resourceType, uint[] resourceId, RuleType type = RuleType.Direct)
            : base(type)
        {
            ResourceType = resourceType;
            ResourceId = resourceId;
        }
    }
    public class ExcludeRule : CollectionRule
    {
        public string[]? CollectionId { get; set; }
        public ExcludeRule(RuleType type = RuleType.Exclude) : base(type) { }
        public ExcludeRule(string[] collectionId, RuleType type) : base(type) { CollectionId = collectionId; }
    }
    public class IncludeRule : CollectionRule
    {
        public string[]? CollectionId { get; set; }
        public IncludeRule(RuleType type = RuleType.Exclude) : base(type) { }
        public IncludeRule(string[] collectionId, RuleType type) : base(type) { CollectionId = collectionId; }
    }
    public class QueryRule : CollectionRule
    {
        public string? QueryExpression { get; set; }
        public QueryRule(RuleType type = RuleType.Exclude) : base(type) { }
        public QueryRule(string queryExpression, string ruleName, RuleType type = RuleType.Exclude)
            : base(type)
        {
            QueryExpression = queryExpression;
            RuleName = ruleName;
        }
    }
    public class Collection
    {
        public string? CollectionId { get; private set; }
        public string Name { get; set; }
        public Lazy<CollectionRule[]>? CollectionRules { get; set; }
        public ResourceType CollectionType { get; set; }
        public CollectionStatus Status { get; private set; }
        public DateTime LastMemberChange { get; private set; }
        public DateTime LastRefresh { get; private set; }
        public Lazy<Collection>? LimitingCollection { get; set; }
        public int MemberCount { get; private set; }
        public bool IsSms { get; private set; }

        public Collection(string name) { Name = name; }

        public static void CreateCollection(string name, string limitingCollectionId, string siteServer)
        {
            ManagementScope scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS");
            scope.Connect();
            ObjectQuery? query = new ObjectQuery("Select Name From __NAMESPACE");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

            string? siteCode = ConfigManagerUtils.Utilities.Console.GetSiteCode(searcher);
            scope = new ManagementScope("\\\\" + siteServer + "\\root\\SMS\\site_" + siteCode);
            scope.Connect();

            if (CheckCollectionName(name, scope) == true) { throw new InvalidOperationException("A collection with the name '" + name + "' already exists."); }
            CheckLimitingCollection(limitingCollectionId, scope);

            ManagementClass smsCollection = new ManagementClass(scope, new ManagementPath("SMS_Collection"), null);
            ManagementObject collection = smsCollection.CreateInstance();
            collection.Properties["Name"].Value = name;
            collection.Properties["LimitToCollectionID"].Value = limitingCollectionId;
            collection.Put();
        }
        public static void CreateCollection(string name, Collection limitingCollection, string siteServer)
        {

        }
        public static void PublishCollection(Collection collection)
        {

        }
        internal static bool CheckCollectionName(string collectionName, ManagementScope scope)
        {
            ObjectQuery query = new ObjectQuery("Select CollectionID from SMS_Collection Where Name = '" + collectionName + "'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            
            if (searcher.Get() is null) { searcher.Dispose(); return false; }
            else { searcher.Dispose(); return true; }
        }
        internal static void CheckLimitingCollection(string collectionId, ManagementScope scope)
        {
            ObjectQuery query = new ObjectQuery("Select CollectionID from SMS_Collection Where CollectionID = '" + collectionId + "'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

            if (searcher.Get() is null) { searcher.Dispose(); }
            else { searcher.Dispose(); throw new InvalidDataException("Limiting collection with ID '" + collectionId + "' not found."); }
        }
    }
}