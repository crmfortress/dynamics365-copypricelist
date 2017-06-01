using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace CRMFortress.CopyPriceList
{
    public class CopyPriceListWorkflow : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // Create a column set to define which attributes should be retrieved.
            ColumnSet attributes = new ColumnSet(true);

            Entity currentPriceLevelEntity = new Entity("pricelevel");
            currentPriceLevelEntity = service.Retrieve("pricelevel", CurrentPriceList.Get<EntityReference>(executionContext).Id, attributes);

            // Create New Price List
            Entity newPriceLevel = new Entity("pricelevel");
            newPriceLevel["name"] = NewPriceListName.Get<string>(executionContext);

            DateTime startDate = NewPriceListStartDate.Get<DateTime>(executionContext);
            newPriceLevel["begindate"] = startDate.Date != DateTime.MinValue ? startDate : DateTime.Now;

            DateTime endDate = NewPriceListEndDate.Get<DateTime>(executionContext);
            newPriceLevel["enddate"] = endDate.Date != DateTime.MinValue ? endDate : DateTime.Now;

            newPriceLevel["transactioncurrencyid"] = currentPriceLevelEntity["transactioncurrencyid"];
            Guid newPriceLevelId = service.Create(newPriceLevel);


            // Get current product price level
            QueryExpression productPriceLevelExpression = new QueryExpression("productpricelevel");

            FilterExpression productPriceLevelFilterExpression = new FilterExpression();
            productPriceLevelFilterExpression.Conditions.Add(new ConditionExpression("pricelevelid", ConditionOperator.Equal, currentPriceLevelEntity["pricelevelid"]));

            productPriceLevelExpression.ColumnSet = new ColumnSet(true);
            productPriceLevelExpression.Criteria = productPriceLevelFilterExpression;

            EntityCollection productPriceLevelList = service.RetrieveMultiple(productPriceLevelExpression);

            // Create new product price level records
            for (int index = 0; productPriceLevelList.Entities != null && index < productPriceLevelList.Entities.Count; index++)
            {
                Entity newProductPriceLevelEntity = new Entity("productpricelevel");
                newProductPriceLevelEntity["pricelevelid"] = new EntityReference("pricelevel", newPriceLevelId);
                newProductPriceLevelEntity["productid"] = productPriceLevelList.Entities[index]["productid"];
                newProductPriceLevelEntity["uomid"] = productPriceLevelList.Entities[index]["uomid"];
                newProductPriceLevelEntity["amount"] = productPriceLevelList.Entities[index]["amount"];
                service.Create(newProductPriceLevelEntity);
            }
        }

        //Define the properties
        [Input("Current Price List")]
        [RequiredArgument]
        [ReferenceTarget("pricelevel")]
        public InArgument<EntityReference> CurrentPriceList { get; set; }

        [RequiredArgument]
        [Input("New Price List : Name")]
        public InArgument<string> NewPriceListName { get; set; }

        [RequiredArgument]
        [Input("New Price List : Start Date")]
        public InArgument<DateTime> NewPriceListStartDate { get; set; }

        [Input("New Price List : End Date")]
        public InArgument<DateTime> NewPriceListEndDate { get; set; }
    }
}
