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
        "Mentor.PlugIns.GetCertificateTemplate", "3f2a6bcd-9e4b-4f2f-8a4d-6a1d8f3e6b79", "", "Mentor.PlugIns (1.0.0.0)", IsolationModeEnum.Sandbox
    )]
    public sealed class GetCertificateTemplate : CodeActivity
    {
        [Input("Course")]
        [ReferenceTarget(men_course.EntityLogicalName)]
        [RequiredArgument]
        public InArgument<EntityReference> CourseId { get; set; }

        [Output("TemplateId")]
        [ReferenceTarget(DocumentTemplate.EntityLogicalName)]
        public OutArgument<EntityReference> TemplateId { get; set; }

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
                    throw new InvalidPluginExecutionException("GetCert: Failed to retrieve tracing service.");
                }
                tracingService.Trace("Entered GetCert.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}", executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId);
                
                context = executionContext.GetExtension<IWorkflowContext>();
                if (context == null)
                {
                    throw new InvalidPluginExecutionException("GetCert: Failed to retrieve workflow context.");
                }
                tracingService.Trace("GetCert.Execute(), Correlation Id: {0}, Initiating User: {1}", context.CorrelationId, context.InitiatingUserId);
                
                tracingService.Trace("GetCert: Organisation Service");
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                service = serviceFactory.CreateOrganizationService(context.UserId);
               
                tracingService.Trace("GetCert: Parameters");
                Guid courseId = CourseId.Get<EntityReference>(executionContext).Id;
                tracingService.Trace("GetCert: Course={0}", courseId.ToString());
                
                tracingService.Trace("CreateCertificate: Retrieving Account");

                Entity course = (Entity)service.Retrieve(men_course.EntityLogicalName, courseId, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));

                // Get Account
                Entity account = (Entity)service.Retrieve(Account.EntityLogicalName, ((EntityReference)course["men_siteaccountid"]).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));

                tracingService.Trace("CreateCertificate: Finding Certificate Template");

                Entity trucktype = null;
                int nontruck = 278460000;
                int noliftheight = 278460000;
                int includeexpirydate = 278460000;
                int notestdate = 290340000;
                if (course.Attributes.Contains("men_trucktypeid"))
                {
                    trucktype = (Entity)service.Retrieve(men_trucktype.EntityLogicalName, ((EntityReference)course["men_trucktypeid"]).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                    if (trucktype.Attributes.Contains("men_nontruck"))
                    {
                        nontruck = ((OptionSetValue)trucktype.Attributes["men_nontruck"]).Value;
                    }
                    if (trucktype.Attributes.Contains("men_noliftheight"))
                    {
                        noliftheight = ((OptionSetValue)trucktype.Attributes["men_noliftheight"]).Value;
                    }
                    if (trucktype.Attributes.Contains("men_includeexpirydate"))
                    {
                        includeexpirydate = ((OptionSetValue)trucktype.Attributes["men_includeexpirydate"]).Value;
                    }
                    if (trucktype.Attributes.Contains("men_notestdate"))
                    {
                        notestdate = ((OptionSetValue)trucktype.Attributes["men_notestdate"]).Value;
                    }
                }

                bool noAttachments = false;
                EntityCollection truckAttachments;
                // If for Training Centre Bookings then get Truck directly from Course. This will always be populated as pre checks are done in previous workflow activities.
                if (course.Contains("men_trainingcentrecoursedetailid"))
                {
                    tracingService.Trace("CreateCertificate: Course is linked to Training Centre Course");
                    // Find if there are any Attachments
                    QueryExpression qeAttachments = new QueryExpression("men_truck");

                    // Add all columns to QEmen_truck.ColumnSet
                    qeAttachments.ColumnSet.AddColumn("men_name");

                    qeAttachments.Criteria.AddCondition("men_truckid", ConditionOperator.Equal, ((EntityReference)course["men_truckid"]).Id);

                    // Add link-entity QEmen_truck_men_instructortruckattachment
                    LinkEntity linkInstructortruckattachment = qeAttachments.AddLink("men_instructortruckattachment", "men_truckid", "men_truckid", JoinOperator.Exists);

                    // Define filter QEmen_truck_men_instructortruckattachment.LinkCriteria
                    linkInstructortruckattachment.LinkCriteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    truckAttachments = service.RetrieveMultiple(qeAttachments);

                    if (truckAttachments.Entities.Count == 0)
                        noAttachments = true;
                }
                else
                {
                    // Find if there are any Attachments
                    QueryExpression qeAttachments = new QueryExpression("men_truck");

                    // Add all columns to QEmen_truck.ColumnSet
                    qeAttachments.ColumnSet.AddColumn("men_name");

                    // Define filter QEmen_truck.Criteria
                    qeAttachments.Criteria.AddCondition("men_courseid", ConditionOperator.Equal, courseId);

                    // Add link-entity QEmen_truck_men_instructortruckattachment
                    LinkEntity linkInstructortruckattachment = qeAttachments.AddLink("men_instructortruckattachment", "men_truckid", "men_truckid", JoinOperator.Exists);

                    // Define filter QEmen_truck_men_instructortruckattachment.LinkCriteria
                    linkInstructortruckattachment.LinkCriteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    truckAttachments = service.RetrieveMultiple(qeAttachments);

                    if (truckAttachments.Entities.Count == 0)
                        noAttachments = true;
                }

                tracingService.Trace("CreateCertificate: Non-Truck = {0}", nontruck.ToString());
                tracingService.Trace("CreateCertificate: No Lift Height = {0} ", noliftheight.ToString());
                tracingService.Trace("CreateCertificate: Include Expiry Date = {0} ", includeexpirydate.ToString());
                tracingService.Trace("CreateCertificate: No Test Date = {0} ", notestdate.ToString());
                tracingService.Trace("CreateCertificate: No Attachments = {0}, Number of Attachments = {1}", noAttachments.ToString(), truckAttachments.Entities.Count.ToString());

                // Second Search men_certificatetemplate for a template that is both for the account and accreditation body
                // If there isn't one that satisfies both the body and account go up the account hierarchy. Once at the top of the hierarchy and there is not a match we use the mentor default for that body
                string docTemplate = GetCertTemplate(service, tracingService, ((EntityReference)course["men_accreditationbodyid"]).Id, account, nontruck, noliftheight, noAttachments, includeexpirydate, notestdate);

                tracingService.Trace("CreateCertificate: docTemplate = {0} ", docTemplate);

                if (docTemplate != string.Empty)
                {
                    this.TemplateId.Set(executionContext, new EntityReference(DocumentTemplate.EntityLogicalName, new Guid(docTemplate)));
                    this.Result.Set(executionContext, true);
                    this.Message.Set(executionContext, docTemplate);
                }
                else
                {
                    this.Result.Set(executionContext, false);
                    this.Message.Set(executionContext, string.Empty);

                }
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
            

            tracingService.Trace("Exiting GetCert.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}", executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId);
        }
        public static string GetCertTemplate(IOrganizationService service, ITracingService tracingService, Guid abody, Entity custacc, int nontruck, int noliftheight, bool noattachment, int includeexpirydate, int notestdate)
        {
            QueryExpression certquery;

            Guid origcustacc = custacc.Id;
            Guid origabody = abody;

            // Check if the account has a broker, if so we set the account to the broker first and 
            // check if there are any Certificates for that broker. 
            // This method calls itself with the broker account instead of the customer account  
            if (custacc.Contains("men_brokerid"))
            {
                Entity broker;
                broker = (Entity)service.Retrieve("men_broker", ((EntityReference)custacc["men_brokerid"]).Id, new ColumnSet(true));
                if (broker.Contains("men_brokeraccountid"))
                {
                    string docTemplate = GetCertTemplate(service, tracingService, abody,
                        (Entity)service.Retrieve(Account.EntityLogicalName, ((EntityReference)broker["men_brokeraccountid"]).Id, new ColumnSet(true)),
                        nontruck, noliftheight, noattachment, includeexpirydate, notestdate);
                    if (docTemplate != string.Empty)
                        return docTemplate;
                }
            }

            while (1 == 1)
            {
                Guid customerAcc = Guid.Empty;
                if (custacc != null && abody != Guid.Empty)
                {
                    customerAcc = custacc.Id;
                    certquery = new QueryExpression
                    {
                        EntityName = "men_certificatetemplate",
                        ColumnSet = new ColumnSet(true),
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
                                        Values = { abody }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_accountid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { customerAcc }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_nontruck",
                                        Operator = ConditionOperator.Equal,
                                        Values = { nontruck }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_noliftheight",
                                        Operator = ConditionOperator.Equal,
                                        Values = { noliftheight }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_noattachment",
                                        Operator = ConditionOperator.Equal,
                                        Values = { noattachment }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_includeexpirydate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { includeexpirydate }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_notestdate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { notestdate }
                                    }
                                }
                            }
                        }
                        }
                    };
                }
                else if (custacc != null && abody == Guid.Empty)
                {
                    customerAcc = custacc.Id;
                    certquery = new QueryExpression
                    {
                        EntityName = "men_certificatetemplate",
                        ColumnSet = new ColumnSet(true),
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
                                        Operator = ConditionOperator.Null
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_accountid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { customerAcc }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_nontruck",
                                        Operator = ConditionOperator.Equal,
                                        Values = { nontruck }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_noliftheight",
                                        Operator = ConditionOperator.Equal,
                                        Values = { noliftheight }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_noattachment",
                                        Operator = ConditionOperator.Equal,
                                        Values = { noattachment }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_includeexpirydate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { includeexpirydate }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "men_notestdate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { notestdate }
                                    }
                                }
                            }
                        }
                        }
                    };
                }
                else if (custacc == null && abody == Guid.Empty)
                {
                    certquery = new QueryExpression
                    {
                        EntityName = "men_certificatetemplate",
                        ColumnSet = new ColumnSet(true),
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
                                            Operator = ConditionOperator.Null
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_accountid",
                                            Operator = ConditionOperator.Null
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_nontruck",
                                            Operator = ConditionOperator.Equal,
                                            Values = { nontruck }
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_noliftheight",
                                            Operator = ConditionOperator.Equal,
                                            Values = { noliftheight }
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_noattachment",
                                            Operator = ConditionOperator.Equal,
                                            Values = { noattachment }
                                        },
                                        new ConditionExpression
                                        {
                                        AttributeName = "men_includeexpirydate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { includeexpirydate }
                                        },
                                        new ConditionExpression
                                        {
                                        AttributeName = "men_notestdate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { notestdate }
                                        }
                                    }
                                }
                            }
                        }
                    };

                }
                else
                {
                    certquery = new QueryExpression
                    {
                        EntityName = "men_certificatetemplate",
                        ColumnSet = new ColumnSet(true),
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
                                            Values = { abody }
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_accountid",
                                            Operator = ConditionOperator.Null
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_nontruck",
                                            Operator = ConditionOperator.Equal,
                                            Values = { nontruck }
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_noliftheight",
                                            Operator = ConditionOperator.Equal,
                                            Values = { noliftheight }
                                        },
                                        new ConditionExpression
                                        {
                                            AttributeName = "men_noattachment",
                                            Operator = ConditionOperator.Equal,
                                            Values = { noattachment }
                                        },
                                        new ConditionExpression
                                        {
                                        AttributeName = "men_includeexpirydate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { includeexpirydate }
                                        },
                                        new ConditionExpression
                                        {
                                        AttributeName = "men_notestdate",
                                        Operator = ConditionOperator.Equal,
                                        Values = { notestdate }
                                        }
                                    }
                                }
                            }
                        }
                    };
                }

                EntityCollection certs = service.RetrieveMultiple(certquery);
                tracingService.Trace("GetCertTemplate: {0}", certs.Entities.Count.ToString());
                if (certs.Entities.Count == 0)
                {
                    // Look at the parent account, if that is populated then retrieve the parent account and re-search for that account
                    if (custacc != null && custacc.Contains("parentaccountid"))
                    {
                        custacc = (Entity)service.Retrieve(Account.EntityLogicalName, ((EntityReference)custacc["parentaccountid"]).Id, new ColumnSet(true));
                    }
                    else
                    {
                        if (custacc == null && abody == Guid.Empty)
                        {
                            return string.Empty;
                        }
                        else if (custacc != null && abody != Guid.Empty)
                        {
                            // Set Accreditation body to Empty so it will look for the Customer General Cert
                            abody = Guid.Empty;
                        }
                        else if (custacc != null && abody == Guid.Empty)
                        {
                            // Set Accreditation body orig 
                            abody = origabody;
                            custacc = null;
                        }
                        else if (custacc == null && abody != Guid.Empty)
                        {
                            // Set Accreditation body to Empty so it will look for the General Cert
                            abody = Guid.Empty;
                        }
                        else
                        {
                            // Set custacc to Null so it will search for a template for the accreditation body not specific to the customer
                            custacc = null;
                        }
                    }

                }
                else if (certs.Entities.Count > 0)
                {
                    return (string)((Entity)certs.Entities[0])["men_templateid"];
                }

            }

        }

    }
}
