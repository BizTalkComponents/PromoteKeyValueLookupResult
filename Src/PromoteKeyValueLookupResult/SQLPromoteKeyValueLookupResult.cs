using BizTalkComponents.Utilities.LookupUtility;
using BizTalkComponents.Utilities.LookupUtility.Repository;
using BizTalkComponents.Utils;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.Message.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using IComponent = Microsoft.BizTalk.Component.Interop.IComponent;

namespace BizTalkComponents.PipelineComponents.PromoteKeyValueLookupResult
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [System.Runtime.InteropServices.Guid("BFF1974C-F57E-41DD-9A21-6241E8825544")]
    [ComponentCategory(CategoryTypes.CATID_Any)]
    public partial class SQLPromoteKeyValueLookupResult : IComponent, IBaseComponent,
                                        IPersistPropertyBag, IComponentUI
    {
        private readonly ILookupRepository _repository = null;

        public SQLPromoteKeyValueLookupResult()
        {
            _repository = new SqlLookupRepository();
        }

        public SQLPromoteKeyValueLookupResult(ILookupRepository repository)
        {
            if (repository == null)
            {
                throw new ArgumentNullException("repository");
            }

            _repository = repository;
        }

        private const string SourcePropertyPathPropertyName = "SourcePropertyPath";
        private const string DestinationPropertyPathPropertyName = "DestinationPropertyPath";
        private const string TableNamePropertyName = "TableName";
        private const string DefaultValuePropertyName = "DefaultValue";
        private const string ThrowIfNotExistsPropertyName = "ThrowIfNotExists";

        [DisplayName("Source Property Path")]
        [Description("The path to the property used as lookup key.")]
        [RequiredRuntime]
        [RegularExpression(@"^.*#.*$",
       ErrorMessage = "A property path should be formatted as namespace#property.")]
        public string SourcePropertyPath { get; set; }

        [DisplayName("Destination Property Path")]
        [Description("The path to the property used to promote to.")]
        [RequiredRuntime]
        [RegularExpression(@"^.*#.*$",
       ErrorMessage = "A property path should be formatted as namespace#property.")]
        public string DestinationPropertyPath { get; set; }

        [DisplayName("Table")]
        [Description("The SQL Table to lookup from.")]
        [RequiredRuntime]
        public string TableName { get; set; }

        [DisplayName("Default value")]
        [Description("Default value to use if no match.")]
        public string DefaultValue { get; set; }

        [DisplayName("ThrowIfNotExists")]
        [Description("Throw exception of key is not found in db.")]
        [RequiredRuntime]
        public bool ThrowIfNotExists { get; set; }

        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            string errorMessage;

            if (!Validate(out errorMessage))
            {
                throw new ArgumentException(errorMessage);
            }

            string lookupKey;

            if (!pInMsg.Context.TryRead(new ContextProperty(SourcePropertyPath), out lookupKey))
            {
                throw new InvalidOperationException("Could not find lookup key in context");
            }

            var util = new LookupUtilityService(_repository);

            var value = util.GetValue(TableName, lookupKey, DefaultValue);

            if (ThrowIfNotExists && string.IsNullOrEmpty(value))
            {
                throw new InvalidOperationException("Could not find value for key " + lookupKey);
            }

            if (!ThrowIfNotExists && string.IsNullOrEmpty(value))
            {
                return pInMsg;
            }

            pInMsg.Context.Promote(new ContextProperty(DestinationPropertyPath), value);

            return pInMsg;
        }

        public void Load(IPropertyBag propertyBag, int errorLog)
        {
            SourcePropertyPath = PropertyBagHelper.ReadPropertyBag(propertyBag, SourcePropertyPathPropertyName, SourcePropertyPath);
            DestinationPropertyPath = PropertyBagHelper.ReadPropertyBag(propertyBag, DestinationPropertyPathPropertyName, DestinationPropertyPath);
            TableName = PropertyBagHelper.ReadPropertyBag(propertyBag, TableNamePropertyName, TableName);
            DefaultValue = PropertyBagHelper.ReadPropertyBag(propertyBag, DefaultValuePropertyName, DefaultValue);
            ThrowIfNotExists = PropertyBagHelper.ReadPropertyBag(propertyBag, ThrowIfNotExistsPropertyName, ThrowIfNotExists);

        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            PropertyBagHelper.WritePropertyBag(propertyBag, SourcePropertyPathPropertyName, SourcePropertyPath);
            PropertyBagHelper.WritePropertyBag(propertyBag, DestinationPropertyPathPropertyName, DestinationPropertyPath);
            PropertyBagHelper.WritePropertyBag(propertyBag, TableNamePropertyName, TableName);
            PropertyBagHelper.WritePropertyBag(propertyBag, DefaultValuePropertyName, DefaultValue);
            PropertyBagHelper.WritePropertyBag(propertyBag, ThrowIfNotExistsPropertyName, ThrowIfNotExists);
        }
    }
}
