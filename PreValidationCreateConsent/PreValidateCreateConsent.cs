using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

 namespace PreValidationCreateConsent
{
    public class PreValidateCreateConsent : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                // בדיקה אם השדות הנדרשים קיימים ואינם ריקים
                if (!entity.Contains("meu_contactid") || entity["meu_contactid"] == null ||
                    !entity.Contains("meu_marketinglistid") || entity["meu_marketinglistid"] == null ||
                    !entity.Contains("meu_consentlevelcode") || entity["meu_consentlevelcode"] == null)
                {
                    throw new InvalidPluginExecutionException("All mandatory fields (Contact, Marketing List, Consent Level) must be filled and not empty.");
                }

            }
        }
    }

}
