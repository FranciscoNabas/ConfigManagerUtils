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
    public class DeviceCollection
    {
        
    }
}