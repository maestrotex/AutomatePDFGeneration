using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuoteTemplateDynamics
{
    public class downloadquotepdf : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") &&
                        context.InputParameters["Target"] is Entity)
            {              
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                Entity preEntity = (Entity)context.PreEntityImages["PreImage"];
                var quotenumber = preEntity.Attributes["quotenumber"];
                var revisionnumber = preEntity.Attributes["revisionnumber"];

                byte[] pdfAsByteArrayRaw = GeneratePDFFromWordTemplate(service, new Guid("EE0BC743-01D4-EB11-BACC-000D3AD87583"),
                    "SOFTCHIEF Quotation", 1084, "quote", entity.Id);

                EntityReference entref = new EntityReference("quote",entity.Id);
                AttachPDFToRecord(service, pdfAsByteArrayRaw, entref, 1084, "Quotation PDF", "The Quote PDF is attached",
                    "QUOTE-"+quotenumber +"-"+revisionnumber);
             }
        }

        public byte[] GeneratePDFFromWordTemplate(IOrganizationService service, Guid? wordTemplateId, string wordTemplateName, int? entityTypeCode, string entityName, Guid entityId)
        {
            // Get the Entity Type code if not known
            if (entityTypeCode == null)
            {
                entityTypeCode = GetObectTypeCodeOfEntity(service, entityName);
            }

            // Get the Word Template ID if not known
            if (wordTemplateId == null)
            {
                wordTemplateId = GetWordTemplateID(service, entityTypeCode, wordTemplateName);
            }

            OrganizationRequest exportPdfAction = new OrganizationRequest("ExportPdfDocument");

            exportPdfAction["EntityTypeCode"] = entityTypeCode;
            exportPdfAction["SelectedTemplate"] = new EntityReference("documenttemplate", (Guid)wordTemplateId);
            exportPdfAction["SelectedRecords"] = "[\'{" + entityId + "}\']";

            OrganizationResponse convertPdfResponse = (OrganizationResponse)service.Execute(exportPdfAction);

            return convertPdfResponse["PdfFile"] as byte[];
        }
        
        private Guid GetWordTemplateID(IOrganizationService service, int? entityTypeCode, string wordTemplateName)
        {
            QueryExpression query = new QueryExpression("documenttemplate");
            query.ColumnSet.AddColumns("name", "associatedentitytypecode");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, wordTemplateName);
            query.Criteria.AddCondition("associatedentitytypecode", ConditionOperator.Equal, (int)entityTypeCode);

            EntityCollection templates = service.RetrieveMultiple(query);

            if (templates.Entities.Count == 0)
            {
                throw new Exception($"No template found with name {wordTemplateName}");
            }
            if (templates.Entities.Count > 1)
            {
                throw new Exception($"More than one template found with name {wordTemplateName}");
            }
            return templates.Entities[0].Id;
        }

        public int GetObectTypeCodeOfEntity(IOrganizationService service, string entityName)
        {
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Entity,
                LogicalName = entityName
            };

            RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
            EntityMetadata AccountEntity = retrieveAccountEntityResponse.EntityMetadata;

            return (int)retrieveAccountEntityResponse.EntityMetadata.ObjectTypeCode;
        }

        public void AttachPDFToRecord(IOrganizationService service, byte[] pdfAsByteArray, 
            EntityReference entityReference, int objectTypeCode, string subject, string noteText, string fileName)
        {
            Entity Annotation = new Entity("annotation");
            Annotation.Attributes["objectid"] = entityReference;
            Annotation.Attributes["objecttypecode"] = objectTypeCode;
            Annotation.Attributes["subject"] = subject;
            Annotation.Attributes["documentbody"] = Convert.ToBase64String(pdfAsByteArray);
            Annotation.Attributes["mimetype"] = @"application/pdf";
            Annotation.Attributes["notetext"] = noteText;
            Annotation.Attributes["filename"] = fileName;
            service.Create(Annotation);
        }
    }

}
