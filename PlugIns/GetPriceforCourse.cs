using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace Mentor.PlugIns
{
    [CrmPluginRegistration(
        "Mentor.PlugIns.GetPriceforCourse", "9d5f2a1b-7c4e-43a9-b8f1-6e2d3c9f8b7a", "", "Mentor.PlugIns (1.0.0.0)", IsolationModeEnum.Sandbox
    )]
    public sealed class GetPriceforCourse : CodeActivity
    {
        [Input("Course")]
        [ReferenceTarget(men_course.EntityLogicalName)]
        [RequiredArgument]
        public InArgument<EntityReference> CourseId { get; set; }

        [Output("Price")]
        public OutArgument<Money> Price { get; set; }

        [Output("Result")]
        public OutArgument<bool> Result { get; set; }

        [Output("Message")]
        public OutArgument<string> Message { get; set; }


        protected override void Execute(CodeActivityContext executionContext)
        {
            IOrganizationService service = null;
            IWorkflowContext context = null;
            ITracingService tracingService = null;

            try
            {
                tracingService = executionContext.GetExtension<ITracingService>();
                if (tracingService == null)
                {
                    throw new InvalidPluginExecutionException("GetPriceforCourse: Failed to retrieve tracing service.");
                }
                tracingService.Trace("Entered GetPriceforCourse.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}", executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId);

                context = executionContext.GetExtension<IWorkflowContext>();
                if (context == null)
                {
                    throw new InvalidPluginExecutionException("GetPriceforCourse: Failed to retrieve workflow context.");
                }
                tracingService.Trace("GetPriceforCourse.Execute(), Correlation Id: {0}, Initiating User: {1}", context.CorrelationId, context.InitiatingUserId);
                tracingService.Trace("GetPriceforCourse: MessageName=" + context.MessageName);
                tracingService.Trace("GetPriceforCourse: Entity=" + context.PrimaryEntityName);
                tracingService.Trace("GetPriceforCourse: Id=" + context.PrimaryEntityId.ToString());
                tracingService.Trace("GetPriceforCourse: Depth=" + context.Depth.ToString());

                tracingService.Trace("GetPriceforCourse: Organisation Service");
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                service = serviceFactory.CreateOrganizationService(context.UserId);

                tracingService.Trace("GetPriceforCourse: Parameters");
                Guid courseId = CourseId.Get<EntityReference>(executionContext).Id;
                tracingService.Trace("GetPriceforCourse: Course={0}", courseId.ToString());
                this.Result.Set(executionContext, false);
                this.Message.Set(executionContext, "");

                Guid pricebookid = Guid.Empty;
                Guid currencyid = Guid.Empty;
                Guid pricelistid = Guid.Empty;
                Guid productid = Guid.Empty;
                Money price = new Money(0);
                Money priceperday = new Money(0);

                ColumnSet coursecols = new ColumnSet(false);
                coursecols.AddColumn("men_name");
                coursecols.AddColumn("men_pricelevelid");
                coursecols.AddColumn("men_pricebookid");
                coursecols.AddColumn("transactioncurrencyid");
                coursecols.AddColumn("men_price");
                coursecols.AddColumn("men_hours");
                coursecols.AddColumn("men_days");
                ColumnSet pricebookcols = new ColumnSet(false);
                pricebookcols.AddColumn("men_name");
                pricebookcols.AddColumn("men_productid");
                pricebookcols.AddColumn("men_standardprice");
                ColumnSet productcols = new ColumnSet(false);
                productcols.AddColumn("name");
                productcols.AddColumn("productnumber");
                productcols.AddColumn("price");
                ColumnSet pricelistcols = new ColumnSet(false);
                pricelistcols.AddColumn("name");
                ColumnSet pricelistitemcols = new ColumnSet(false);
                pricelistitemcols.AddColumn("amount");


                tracingService.Trace("GetPriceforCourse: Retrieve Course");
                men_course course = (men_course)service.Retrieve(men_course.EntityLogicalName, courseId, coursecols);

                tracingService.Trace("GetPriceforCourse: Days");
                decimal days = new decimal(1.0);
                if (course.Contains("men_days"))
                {
                    if (course.men_days != null)
                    {
                        if (course.men_days.Value > 0)
                        {
                            tracingService.Trace("GetPriceforCourse: Days={0}", course.men_days.Value.ToString());
                            days = (int)course.men_days.Value;
                            tracingService.Trace("GetPriceforCourse: Days={0}", days.ToString());
                        }
                        if (days < 1)
                        {
                            days = 1;
                            tracingService.Trace("GetPriceforCourse: Days={0}", days.ToString());
                        }
                    }
                }

                tracingService.Trace("GetPriceforCourse: Price List");
                if (pricelistid == Guid.Empty)
                {
                    if (course.Contains("men_pricelevelid"))
                    {
                        if (course.men_pricelevelid != null)
                        {
                            pricelistid = course.men_pricelevelid.Id;
                        }
                    }
                }
                if (pricelistid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceforCourse: Price List={0}", pricelistid.ToString());
                    tracingService.Trace("GetPriceforCourse: Price List={0}", course.men_pricelevelid.Name);
                }
                else
                {
                    tracingService.Trace("GetPriceforCourse: No Price List");

                    tracingService.Trace("GetPriceforCourse: Mentor Settings");
                    QueryExpression settingsquery = new QueryExpression
                    {
                        EntityName = men_mentorsettings.EntityLogicalName,
                        ColumnSet = new ColumnSet(true)
                    };
                    settingsquery.Criteria.AddCondition("statecode", ConditionOperator.Equal, (int)men_mentorsettingsState.Active);
                    EntityCollection settings = service.RetrieveMultiple(settingsquery);
                    tracingService.Trace("GetPriceforCourse: Count={0}", settings.Entities.Count.ToString());
                    if (settings.Entities.Count != 1)
                    {
                        tracingService.Trace("GetPriceforCourse: Mentor Settings not found");
                    }
                    else
                    {
                        men_mentorsettings setting = (men_mentorsettings)settings[0];
                        tracingService.Trace("GetPriceforCourse: Mentor Settings {0}", setting.men_name);
                        if (setting.Contains("men_pricelevelid"))
                        {
                            if (setting.men_defaultpricelevelid != null)
                            {
                                pricelistid = setting.men_defaultpricelevelid.Id;
                                tracingService.Trace("GetPriceforCourse: Price List={0}", pricelistid.ToString());
                                tracingService.Trace("GetPriceforCourse: Price List={0}", setting.men_defaultpricelevelid.Name);
                            }
                        }
                    }


                    if (pricelistid == Guid.Empty)
                    {
                        this.Result.Set(executionContext, false);
                        this.Message.Set(executionContext, "No Price List");
                        return;
                    }
                }
                if (currencyid == Guid.Empty)
                {
                    if (course.Contains("transactioncurrencyid"))
                    {
                        if (course.TransactionCurrencyId != null)
                        {
                            currencyid = course.TransactionCurrencyId.Id;
                            tracingService.Trace("GetPriceforCourse: Currency={0}", currencyid.ToString());
                        }
                    }
                }
                if (currencyid == Guid.Empty)
                {
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "No Currency");
                    return;
                }


                tracingService.Trace("GetPriceforCourse: Price Book");
                if (pricebookid == Guid.Empty)
                {
                    if (course.Contains("men_pricebookid"))
                    {
                        if (course.men_pricebookid != null)
                        {
                            pricebookid = course.men_pricebookid.Id;
                        }
                    }
                }
                if (pricebookid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceforCourse: Price Book={0}", pricebookid.ToString());
                    tracingService.Trace("GetPriceforCourse: Price Book={0}", course.men_pricebookid.Name);
                }
                else
                {
                    tracingService.Trace("GetPriceforCourse: No Price Book");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "No Price Book");
                    return;
                }

                tracingService.Trace("GetPriceforCourse: Retrieve Price Book");
                men_pricebook pricebook = (men_pricebook)service.Retrieve(men_pricebook.EntityLogicalName, pricebookid, pricebookcols);
                tracingService.Trace("GetPriceforCourse: Price Book={0}", pricebook.men_name);
                if (pricebook.Contains("men_productid"))
                {
                    if (pricebook.men_productid != null)
                    {
                        productid = pricebook.men_productid.Id;
                        tracingService.Trace("GetPriceforCourse: Price Book Product={0}", productid.ToString());
                        tracingService.Trace("GetPriceforCourse: Price Book Product={0}", pricebook.men_productid.Name);
                    }
                }


                tracingService.Trace("GetPriceforCourse: Product");
                if (productid == Guid.Empty)
                {
                    if (pricebook.Contains("men_productid"))
                    {
                        if (pricebook.men_productid != null)
                        {
                            productid = pricebook.men_productid.Id;
                        }
                    }
                }
                if (productid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceforCourse: Product={0}", productid.ToString());
                    tracingService.Trace("GetPriceforCourse: Product={0}", pricebook.men_productid.Name);
                }
                else
                {
                    tracingService.Trace("GetPriceforCourse: No Product");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "No Product");
                    return;
                }


                tracingService.Trace("GetPriceforCourse: Retrieve Product");
                Product product = (Product)service.Retrieve(Product.EntityLogicalName, productid, productcols);
                tracingService.Trace("GetPriceforCourse: Product={0}", product.Name);
                tracingService.Trace("GetPriceforCourse: Product={0}", product.ProductNumber);
                if (product.Contains("price"))
                {
                    if (product.Price != null)
                    {
                        if (product.Price.Value > 0)
                        {
                            priceperday = product.Price;
                            tracingService.Trace("GetPriceforCourse: Product List Price Per Day={0}", priceperday.Value.ToString());
                        }
                    }
                }
                if (pricebook.Contains("men_standardprice"))
                {
                    if (pricebook.men_standardprice != null)
                    {
                        if (pricebook.men_standardprice.Value > 0)
                        {
                            price = pricebook.men_standardprice;
                            tracingService.Trace("GetPriceforCourse: Price Book Standard Price={0}", price.Value.ToString());
                        }
                    }
                }


                tracingService.Trace("GetPriceforCourse: Retrieve Price List");
                PriceLevel pricelist = (PriceLevel)service.Retrieve(PriceLevel.EntityLogicalName, pricelistid, pricelistcols);
                tracingService.Trace("GetPriceforCourse: Price List={0}", pricelist.Name);


                tracingService.Trace("GetPriceforCourse: Price List Item");

                QueryExpression pricelistitemquery = new QueryExpression
                {
                    EntityName = ProductPriceLevel.EntityLogicalName,
                    ColumnSet = pricelistitemcols,
                    Criteria = new FilterExpression
                    {
                        FilterOperator = LogicalOperator.And,
                        Filters =
                        {
                            new FilterExpression
                            {
                                FilterOperator = LogicalOperator.And,
                                Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "pricelevelid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { pricelistid }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "productid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { productid }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "transactioncurrencyid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { currencyid }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "pricingmethodcode",
                                        Operator = ConditionOperator.Equal,
                                        Values = { (int)productpricelevel_pricingmethodcode.CurrencyAmount }
                                    }
                                }
                            }
                        }
                    }
                };


                //QueryExpressionToFetchXmlRequest pricelisitemrequest = new QueryExpressionToFetchXmlRequest()
                //{
                //    Query = pricelistitemquery
                //};
                //QueryExpressionToFetchXmlResponse pricelistitemresp = (QueryExpressionToFetchXmlResponse)service.Execute(pricelisitemrequest);
                //tracingService.Trace("GetPriceforCourse: Query=" + pricelistitemresp.FetchXml);


                EntityCollection pricelistitems = service.RetrieveMultiple(pricelistitemquery);
                tracingService.Trace("GetPriceforCourse: Price List Items={0}", pricelistitems.Entities.Count.ToString());
                if (pricelistitems.Entities.Count == 0)
                {
                    tracingService.Trace("GetPriceforCourse: Price List Item not found");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "Price List Item not found");
                    return;
                }
                if (pricelistitems.Entities.Count > 1)
                {
                    tracingService.Trace("GetPriceforCourse: More than 1 Price List Item found");
                }
                ProductPriceLevel item = (ProductPriceLevel)pricelistitems[0];
                if (item.Contains("amount"))
                {
                    if (item.Amount != null)
                    {
                        if (item.Amount.Value > 0)
                        {
                            priceperday = item.Amount;
                            tracingService.Trace("GetPriceforCourse: Product Price List Item Price Per Day={0}", priceperday.Value.ToString());
                        }
                    }
                }


                if (priceperday.Value > 0)
                {
                    tracingService.Trace("GetPriceforCourse: Price Per Day={0}", priceperday.Value.ToString());
                    price = new Money(priceperday.Value * days);
                }
                tracingService.Trace("GetPriceforCourse: Price={0}", price.Value.ToString());
                this.Price.Set(executionContext, price);
                this.Result.Set(executionContext, true);
                tracingService.Trace("GetPriceforCourse: Completed");

            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace(string.Format("Error in {0}\r\n{1}", MethodInfo.GetCurrentMethod().Name, ex.Message));
                throw new InvalidPluginExecutionException(string.Format("Error in {0}\r\n{1}", MethodInfo.GetCurrentMethod().Name, ex.Message));
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("Error in {0}\r\n{1}", MethodInfo.GetCurrentMethod().Name, ex.Message));
                throw new InvalidPluginExecutionException(string.Format("Error in {0}\r\n{1}", MethodInfo.GetCurrentMethod().Name, ex.Message));
            }


            tracingService.Trace("Exiting GetPriceforCourse.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}", executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId);
        }
    }
}
