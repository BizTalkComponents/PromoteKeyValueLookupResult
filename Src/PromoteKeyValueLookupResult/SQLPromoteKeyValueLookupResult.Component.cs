using BizTalkComponents.Utilities.LookupUtility;
using BizTalkComponents.Utilities.LookupUtility.Repository;
using BizTalkComponents.Utils;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.Message.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BizTalkComponents.PipelineComponents.PromoteKeyValueLookupResult
{
    
    public partial class SQLPromoteKeyValueLookupResult 
    {
        public string Name { get { return "SQLPromoteKeyValueLookupResult"; } }
        public string Version { get { return "1.0"; } }
        public string Description { get { return "Will promote result from a key value store lookup"; } }

        public void GetClassID(out Guid classID)
        {
            classID = new Guid("EB9816E2-86B1-47D5-8CBE-8602773CCBD6");
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
