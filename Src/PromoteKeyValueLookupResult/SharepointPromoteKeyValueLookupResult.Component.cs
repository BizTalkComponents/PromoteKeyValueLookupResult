using System;
using System.Collections;
using System.Linq;
using BizTalkComponents.Utils;

namespace BizTalkComponents.PipelineComponents.PromoteKeyValueLookupResult
{
    public partial class SharepointPromoteKeyValueLookupResult
    {
        public string Name { get { return "SharePointPromoteKeyValueLookupResult"; } }
        public string Version { get { return "1.0"; } }
        public string Description { get { return "Will promote result from a key value store lookup"; } }

        public void GetClassID(out Guid classID)
        {
            classID = new Guid("74E4FE33-872A-11E7-AFE1-61C8FF21A5EE");
        }

        public void InitNew()
        {

        }

        public IEnumerator Validate(object projectSystem)
        {
            return ValidationHelper.Validate(this, false).ToArray().GetEnumerator();
        }

        public bool Validate(out string errorMessage)
        {
            var errors = ValidationHelper.Validate(this, true).ToArray();

            if (errors.Any())
            {
                errorMessage = string.Join(",", errors);

                return false;
            }

            errorMessage = string.Empty;

            return true;
        }

        public IntPtr Icon { get { return IntPtr.Zero; } }
    }
}
