using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using BizTalkComponents.Utils;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.Message.Interop;
using IComponent = Microsoft.BizTalk.Component.Interop.IComponent;
using BizTalkComponents.Utilities.LookupUtility.Repository;
using BizTalkComponents.Utilities.LookupUtility;

namespace BizTalkComponents.PipelineComponents.PromoteKeyValueLookupResult
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [System.Runtime.InteropServices.Guid("74E4FE30-872A-11E7-AFE1-61C8FF21A5EE")]
    [ComponentCategory(CategoryTypes.CATID_Any)]
    public partial class SharepointPromoteKeyValueLookupResult : IComponent, IBaseComponent,
                                        IPersistPropertyBag, IComponentUI
    {
        private readonly ILookupRepository _repository = null;

        public SharepointPromoteKeyValueLookupResult()
        {
            _repository = new SharepointLookupRepository();
        }

        public SharepointPromoteKeyValueLookupResult(ILookupRepository repository)
        {
            if (repository == null)
            {
                throw new ArgumentNullException("repository");
            }

            _repository = repository;
        }

        private const string SourcePropertyPathPropertyName = "SourcePropertyPath";
        private const string DestinationPropertyPathPropertyName = "DestinationPropertyPath";
        private const string ListNamePropertyName = "ListName";
        private const string DefaultValuePropertyName = "DefaultValue";


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

        [DisplayName("List name")]
        [Description("The Sharepoint list to lookup from.")]
        [RequiredRuntime]
        public string ListName { get; set; }

        [DisplayName("Default value")]
        [Description("Default value to use if no match.")]
        public string DefaultValue { get; set; }

        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            string errorMessage;

            if (!Validate(out errorMessage))
            {
                throw new ArgumentException(errorMessage);
            }

            string lookupKey; 
            
            if(!pInMsg.Context.TryRead(new ContextProperty(SourcePropertyPath), out lookupKey))
            {
                throw new InvalidOperationException("Could not find lookup key in context");
            }

            var util = new LookupUtilityService(_repository);

            var value = util.GetValue(ListName, lookupKey,DefaultValue);

            if (string.IsNullOrEmpty(value))
            {
                throw new InvalidOperationException("Could not find value for key " + lookupKey);
            }

            pInMsg.Context.Promote(new ContextProperty(DestinationPropertyPath), value);

            return pInMsg;
        }

        public void Load(IPropertyBag propertyBag, int errorLog)
        {
            SourcePropertyPath = PropertyBagHelper.ReadPropertyBag(propertyBag, SourcePropertyPathPropertyName, SourcePropertyPath);
            DestinationPropertyPath = PropertyBagHelper.ReadPropertyBag(propertyBag, DestinationPropertyPathPropertyName, DestinationPropertyPath);
            ListName = PropertyBagHelper.ReadPropertyBag(propertyBag, ListNamePropertyName, ListName);
            DefaultValue = PropertyBagHelper.ReadPropertyBag(propertyBag, DefaultValuePropertyName, DefaultValue);

        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            PropertyBagHelper.WritePropertyBag(propertyBag, SourcePropertyPathPropertyName, SourcePropertyPath);
            PropertyBagHelper.WritePropertyBag(propertyBag, DestinationPropertyPathPropertyName, DestinationPropertyPath);
            PropertyBagHelper.WritePropertyBag(propertyBag, ListNamePropertyName, ListName);
            PropertyBagHelper.WritePropertyBag(propertyBag, DefaultValuePropertyName, DefaultValue);

        }
    }
}
