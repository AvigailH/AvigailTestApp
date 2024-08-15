using System;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Web.Services.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace PreCreateConsent
{
    public class PreCreateConsent : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                if (entity.Contains("meu_consentlevelcode") && entity.Contains("meu_contactid") && entity.Contains("meu_marketinglistid"))
                {
                    int consentLevel = entity.GetAttributeValue<OptionSetValue>("meu_consentlevelcode").Value;
                    Guid contactId = entity.GetAttributeValue<EntityReference>("meu_contactid").Id;
                    Guid marketingListId = entity.GetAttributeValue<EntityReference>("meu_marketinglistid").Id;

                    if (consentLevel == 1) // "מסכים"
                    {
                        // הוספת איש קשר לרשימת שיווק
                        AddContactToMarketingList(service, contactId, marketingListId);
                    }
                    else
                    {
                        // הסרת איש קשר מרשימת שיווק
                        RemoveContactFromMarketingList(service, contactId, marketingListId);
                    }
                }
            }
        }

        private void AddContactToMarketingList(IOrganizationService service, Guid contactId, Guid marketingListId)
        {
            // הוספת איש קשר לרשימת שיווק
            Entity entity = new Entity("listmember");
            entity["listid"] = new EntityReference("list", marketingListId);
            entity["entityid"] = new EntityReference("contact", contactId);
            service.Create(entity);

            /*CreateRequest createRequest = new CreateRequest
            {
                Target = entity
            };

            CreateResponse response = (CreateResponse)service.Execute(createRequest);
            Guid newRecordId = response.id;*/

        }

        private void RemoveContactFromMarketingList(IOrganizationService service, Guid contactId, Guid marketingListId)
        {
            // הסרת איש קשר מרשימת שיווק
            QueryExpression query = new QueryExpression("listmember")
            {
                Criteria = new FilterExpression
                {
                    Conditions =
            {
                new ConditionExpression("listid", ConditionOperator.Equal, marketingListId),
                new ConditionExpression("entityid", ConditionOperator.Equal, contactId)
            }
                }
            };

            Entity listMember = service.RetrieveMultiple(query).Entities.FirstOrDefault();
            if (listMember != null)
            {
                service.Delete("listmember", listMember.Id);
            }
        }

    }

}
