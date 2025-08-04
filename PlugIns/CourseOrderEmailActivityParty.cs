using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.ServiceModel;

namespace Mentor.PlugIns
{
    [CrmPluginRegistration(
    "Update",
    Email.EntityLogicalName,
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "men_generateactivityparties",
    "Mentor.PlugIns.CourseOrderEmailActivityParty: Update of email",
    1000,
    IsolationModeEnum.Sandbox
    , Image1Type = ImageTypeEnum.PostImage
    , Image1Name = "PostImage"
    , Image1Attributes = "men_generateactivityparties, regardingobjectid"
)]
    public class CourseOrderEmailActivityParty : PluginBase
    {
        protected override void ExecuteCDSPlugin(LocalPluginContext localcontext)
        {
            var service = localcontext.OrganizationService;
            var tracingService = localcontext.TracingService;
            tracingService.Trace("CourseOrderEmailActivityParty: Start");
            try
            {
                //context
                var emailid = localcontext.PluginExecutionContext.PrimaryEntityId;
                var emailTarget = localcontext.Target.ToEntity<Email>();
                var emailEntity = localcontext.MergedPostTarget.ToEntity<Email>();               

               //check contents of postimage                
                if (!emailEntity.Contains("men_generateactivityparties"))
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: No Generate Trigger in Context/PostImage");
                    return;
                }
                if (emailEntity.men_generateactivityparties == null)
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: Generate Trigger is null in Context/PostImage");
                    return;
                }
                if (emailEntity.men_generateactivityparties.Value != true)
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: Generate Order Trigger is Off in Context/PostImage");
                    return;
                }
                if (!emailEntity.Contains("regardingobjectid"))
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: No Regarding in Context/PostImage");
                    return;
                }
                if (emailEntity.RegardingObjectId == null)
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: Regarding is null in Context/PostImage");
                    return;
                }
                if (emailEntity.RegardingObjectId.Id == null)
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: Regarding Id is null in Context/PostImage");
                    return;
                }
                if (emailEntity.RegardingObjectId.LogicalName != men_courseorder.EntityLogicalName)
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: Regarding is not Course Order in Context/PostImage");
                    return;
                }

                //course order
                tracingService.Trace("CourseOrderEmailActivityParty: Course Order {0}", emailEntity.RegardingObjectId.Id.ToString());
                ColumnSet courseordercols = new ColumnSet(false);
                courseordercols.AddColumn("men_contactid");
                courseordercols.AddColumn("ownerid");
                men_courseorder courseorder = (men_courseorder)service.Retrieve(men_courseorder.EntityLogicalName, emailEntity.RegardingObjectId.Id, courseordercols);
                
                Email updateemail = new Email
                {
                    Id = emailid
                };

                //from
                if (courseorder.Contains("ownerid"))
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: From {0}", courseorder.OwnerId.Id.ToString());
                    ActivityParty fromParty = new ActivityParty
                    {
                        PartyId = new EntityReference(SystemUser.EntityLogicalName, courseorder.OwnerId.Id)
                    };
                    updateemail.From = new ActivityParty[] { fromParty };
                }

                //to
                if (courseorder.Contains("men_contactid"))
                {
                    tracingService.Trace("CourseOrderEmailActivityParty: To {0}", courseorder.men_contactid.Id.ToString());
                    ActivityParty toParty = new ActivityParty
                    {
                        PartyId = new EntityReference(Contact.EntityLogicalName, courseorder.men_contactid.Id)
                    };
                    updateemail.To = new ActivityParty[] { toParty };
                }

                //add activityparties to email
                service.Update(updateemail);
                         
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace("CourseOrderEmailActivityParty: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the CourseOrderEmailActivityParty plugin " + ex.Message, ex);
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace("CourseOrderEmailActivityParty: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("A standard exception error occurred in the CourseOrderEmailActivityParty plugin " + ex.Message, ex);
            }
            tracingService.Trace("CourseOrderEmailActivityParty: End");
        }
    }
}
