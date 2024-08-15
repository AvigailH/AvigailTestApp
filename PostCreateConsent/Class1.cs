using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace PostCreateConsent
{
    public class CancelPreviousConsents : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity newConsent = (Entity)context.InputParameters["Target"];

                if (newConsent.Contains("meu_contactid") && newConsent.Contains("meu_marketinglistid"))
                {
                    Guid contactId = newConsent.GetAttributeValue<EntityReference>("meu_contactid").Id;
                    Guid marketingListId = newConsent.GetAttributeValue<EntityReference>("meu_marketinglistid").Id;

                    // חיפוש וביטול הסכמות קודמות
                    CancelExistingConsents(service, contactId, marketingListId);
                }
            }
        }

        private void CancelExistingConsents(IOrganizationService service, Guid contactId, Guid marketingListId)
        {
            QueryExpression query = new QueryExpression("meu_avigailconsent")
            {
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("meu_contactid", ConditionOperator.Equal, contactId),
                        new ConditionExpression("meu_marketinglistid", ConditionOperator.Equal, marketingListId),
                        new ConditionExpression("meu_ConsentLevelCode", ConditionOperator.Equal, 1) // מצב פעיל
                    }
                }
            };

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count > 0)
            {
                // יצירת בקשה לעדכון מרובה
                ExecuteMultipleRequest executeMultipleRequest = new ExecuteMultipleRequest()
                {
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = false,
                        ReturnResponses = false
                    },
                    Requests = new OrganizationRequestCollection()
                };

                foreach (Entity consent in results.Entities)
                {
                    // שינוי מצב ההסכמה ללא פעיל
                    consent["statecode"] = new OptionSetValue(0); // מצב לא פעיל

                    // הוספת העדכון לבקשת העדכונים המרובה
                    UpdateRequest updateRequest = new UpdateRequest { Target = consent };
                    executeMultipleRequest.Requests.Add(updateRequest);
                }

                // ביצוע העדכונים בפעם אחת
                service.Execute(executeMultipleRequest);
            }
        }
    }
}
