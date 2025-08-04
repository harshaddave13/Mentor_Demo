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
        "Mentor.PlugIns.GetPriceBook", "bb7f3e3d-2c51-4e1f-b5f9-9d6a2e9b7d15", "", "Mentor.PlugIns (1.0.0.0)", IsolationModeEnum.Sandbox
    )]
    public sealed class GetPriceBook : CodeActivity
    {
        [Input("Course")]
        [ReferenceTarget(men_course.EntityLogicalName)]
        [RequiredArgument]
        public InArgument<EntityReference> CourseId { get; set; }

        [Output("Price Book")]
        [ReferenceTarget(men_pricebook.EntityLogicalName)]
        public OutArgument<EntityReference> PriceBookId { get; set; }

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
                    throw new InvalidPluginExecutionException("GetPriceBook: Failed to retrieve tracing service.");
                }
                tracingService.Trace("Entered GetPriceBook.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}", executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId);
                
                context = executionContext.GetExtension<IWorkflowContext>();
                if (context == null)
                {
                    throw new InvalidPluginExecutionException("GetPriceBook: Failed to retrieve workflow context.");
                }
                tracingService.Trace("GetPriceBook.Execute(), Correlation Id: {0}, Initiating User: {1}", context.CorrelationId, context.InitiatingUserId);
                tracingService.Trace("GetPriceBook: MessageName=" + context.MessageName);
                tracingService.Trace("GetPriceBook: Entity=" + context.PrimaryEntityName);
                tracingService.Trace("GetPriceBook: Id=" + context.PrimaryEntityId.ToString());
                tracingService.Trace("GetPriceBook: Depth=" + context.Depth.ToString());
                
                tracingService.Trace("GetPriceBook: Organisation Service");
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                service = serviceFactory.CreateOrganizationService(context.UserId);
                
                tracingService.Trace("GetPriceBook: Parameters");
                Guid courseId = CourseId.Get<EntityReference>(executionContext).Id;
                tracingService.Trace("GetPriceBook: Course={0}", courseId.ToString());
                this.Result.Set(executionContext, false);
                this.Message.Set(executionContext, "");
                
                Guid accreditationbodyid = Guid.Empty;
                Guid truckcategoryid = Guid.Empty;
                Guid trucktypeid = Guid.Empty;
                Guid experiencelevelid = Guid.Empty;
                Guid attachmentid = Guid.Empty;
                int numberofdelegates = 0;
                Guid pricebookid = Guid.Empty;
                
                ColumnSet coursecols = new ColumnSet(false);
                coursecols.AddColumn("men_accreditationbodyid");
                coursecols.AddColumn("men_truckcategoryid");
                coursecols.AddColumn("men_trucktypeid");
                coursecols.AddColumn("men_experiencelevelid");
                coursecols.AddColumn("men_attachmentid");
                coursecols.AddColumn("men_numberofdelegates");
                ColumnSet pricebookcols = new ColumnSet(false);
                
                tracingService.Trace("GetPriceBook: Retrieve Course");
                men_course course = (men_course)service.Retrieve(men_course.EntityLogicalName, courseId, coursecols);
                
                tracingService.Trace("GetPriceBook: Accreditation Body");
                if (accreditationbodyid == Guid.Empty)
                {
                    if (course.Contains("men_accreditationbodyid"))
                    {
                        if (course.men_accreditationbodyid != null)
                        {
                            accreditationbodyid = course.men_accreditationbodyid.Id;
                        }
                    }
                }
                if (accreditationbodyid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceBook: Accreditation Body={0}", accreditationbodyid.ToString());
                }
                else
                {
                    tracingService.Trace("GetPriceBook: No Accreditation Body");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "No Accreditation Body");
                    return;
                }
                
                tracingService.Trace("GetPriceBook: Truck Category");
                if (truckcategoryid == Guid.Empty)
                {
                    if (course.Contains("men_truckcategoryid"))
                    {
                        if (course.men_truckcategoryId != null)
                        {
                            truckcategoryid = course.men_truckcategoryId.Id;
                        }
                    }
                }
                if (truckcategoryid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceBook: Truck Category={0}", truckcategoryid.ToString());
                }
                else
                {
                    tracingService.Trace("GetPriceBook: No Truck Category");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "No Truck Category");
                    return;
                }
                
                tracingService.Trace("GetPriceBook: Truck Type");
                if (trucktypeid == Guid.Empty)
                {
                    if (course.Contains("men_trucktypeid"))
                    {
                        if (course.men_trucktypeId != null)
                        {
                            trucktypeid = course.men_trucktypeId.Id;
                        }
                    }
                }
                if (trucktypeid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceBook: Truck Type={0}", trucktypeid.ToString());
                }
                else
                {
                    tracingService.Trace("GetPriceBook: No Truck Type");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "No Truck Type");
                    return;
                }
                
                tracingService.Trace("GetPriceBook: Experience Level");
                if (experiencelevelid == Guid.Empty)
                {
                    if (course.Contains("men_experiencelevelid"))
                    {
                        if (course.men_experiencelevelid != null)
                        {
                            experiencelevelid = course.men_experiencelevelid.Id;
                        }
                    }
                }
                if (experiencelevelid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceBook: Experience Level={0}", experiencelevelid.ToString());
                }
                else
                {
                    tracingService.Trace("GetPriceBook: No Experience Level");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "No Experience Level");
                    return;
                }
                
                tracingService.Trace("GetPriceBook: Attachment");
                if (attachmentid == Guid.Empty)
                {
                    if (course.Contains("men_attachmentid"))
                    {
                        if (course.men_attachmentid != null)
                        {
                            attachmentid = course.men_attachmentid.Id;
                        }
                    }
                }
                if (attachmentid != Guid.Empty)
                {
                    tracingService.Trace("GetPriceBook: Attachment={0}", attachmentid.ToString());
                }
                else
                {
                    tracingService.Trace("GetPriceBook: No Attachment");
                }
                
                if (course.Contains("men_numberofdelegates"))
                {
                    if (course.men_numberofdelegates != null)
                    {
                        numberofdelegates = course.men_numberofdelegates.Value;
                    }
                }
                tracingService.Trace("GetPriceBook: Number of Delegates={0}", numberofdelegates.ToString());
                
                tracingService.Trace("GetPriceBook: Price Book");

                
                QueryExpression pricebookquery = new QueryExpression
                {
                    EntityName = men_pricebook.EntityLogicalName,
                    ColumnSet = pricebookcols,
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
                                        AttributeName = "men_accreditationbodyid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { accreditationbodyid }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_trucktypeid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { trucktypeid }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_experiencelevelid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { experiencelevelid }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_numberofdelegates",
                                        Operator = ConditionOperator.Equal,
                                        Values = { numberofdelegates }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "statecode",
                                        Operator = ConditionOperator.Equal,
                                        Values = { (int)men_pricebookState.Active }
                                    }
                                }
                            }
                        }
                    }
                };
                //if (attachmentid != Guid.Empty)
                //{
                //    pricebookquery.Criteria.AddCondition("men_attachmentid", ConditionOperator.Equal, attachmentid);
                //}
                
                //QueryExpressionToFetchXmlRequest pricebookrequest = new QueryExpressionToFetchXmlRequest()
                //{
                //    Query = priceboookquery
                //};
                //QueryExpressionToFetchXmlResponse pricebookresp = (QueryExpressionToFetchXmlResponse)service.Execute(pricebookrequest);
                //tracingService.Trace("GetPriceBook: Query=" + pricebookresp.FetchXml);
               
                EntityCollection pricebooks = service.RetrieveMultiple(pricebookquery);
                tracingService.Trace("GetPriceBook: Price Books={0}", pricebooks.Entities.Count.ToString());
                if (pricebooks.Entities.Count == 0)
                {
                    tracingService.Trace("GetPriceBook: Price Book not found");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "Price Book not found");
                    return;
                }
                if (pricebooks.Entities.Count > 1)
                {
                    tracingService.Trace("GetPriceBook: More than 1 Price Book found");
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "More than 1 Price Book found");
                    return;
                }
                pricebookid = pricebooks[0].Id;                

                
                if (pricebookid != Guid.Empty)
                {
                    this.PriceBookId.Set(executionContext, new EntityReference(men_pricebook.EntityLogicalName, pricebookid));
                    this.Result.Set(executionContext, true);
                    this.Message.Set(executionContext, "Price Book found");
                }
                else
                {
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, "Price Book not found");
                }
                tracingService.Trace("GetPriceBook: Completed");
               
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
            

            tracingService.Trace("Exiting GetPriceBook.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}", executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId);
        }
    }

}
