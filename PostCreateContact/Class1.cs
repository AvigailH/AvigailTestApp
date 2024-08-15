using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace PostCreateContact
{
    public class AddContactToDefaultMarketingLists : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // קבלת ההקשר והשירות מה- Service Provider
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity contact = (Entity)context.InputParameters["Target"];

                // השגת משתנה סביבה עם קודי רשימות שיווק
                string marketingListCodes = GetEnvironmentVariable(service, "meu_AvigailConsentDefault");

                if (!string.IsNullOrEmpty(marketingListCodes))
                {
                    var codesArray = marketingListCodes.Split(',');

                    foreach (var code in codesArray)
                    {
                        // המרה ל-int ובדיקה שההמרה הצליחה
                        if (int.TryParse(code, out int marketingListCodeInt))
                        {
                            // הוספת איש קשר לרשימת שיווק לפי הקוד מהמשתנה
                            AddContactToMarketingList(service, contact.Id, marketingListCodeInt);
                        }
                    }
                }
            }
        }

        private string GetEnvironmentVariable(IOrganizationService service, string variableName)
        {
            // הבאת ערך משתנה סביבה
            QueryExpression query = new QueryExpression("environmentvariabledefinition")
            {
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("schemaname", ConditionOperator.Equal, variableName)
                    }
                },
                ColumnSet = new ColumnSet("defaultvalue")
            };

            Entity envVarDef = service.RetrieveMultiple(query).Entities.FirstOrDefault();

            return envVarDef != null ? envVarDef.GetAttributeValue<string>("defaultvalue") : string.Empty;
        }

        private void AddContactToMarketingList(IOrganizationService service, Guid contactId, int marketingListCodeInt)
        {
            // מציאת רשימת שיווק לפי קוד
            QueryExpression query = new QueryExpression("list")
            {
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("meu_codeint", ConditionOperator.Equal, marketingListCodeInt)
                    }
                }
            };

            // שימוש ב-RetrieveMultiple כדי למצוא רשימת שיווק אחת בלבד
            Entity marketingList = service.RetrieveMultiple(query).Entities.FirstOrDefault();

            if (marketingList != null)
            {
                Guid marketingListId = marketingList.Id;

                // הוספת איש קשר לרשימת השיווק
                Entity entity = new Entity("listmember");
                entity["listid"] = new EntityReference("list", marketingListId);
                entity["entityid"] = new EntityReference("contact", contactId);
                service.Create(entity);
            }
        }
    }
}
