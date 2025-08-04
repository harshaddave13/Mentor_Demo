using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.ServiceModel;

namespace Mentor.PlugIns
{
    [CrmPluginRegistration(
    "Create",
    men_coursecostline.EntityLogicalName,
    StageEnum.PreOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "Mentor.PlugIns.CourseCostLineSequence: Create of men_coursecostline",
    1000,
    IsolationModeEnum.Sandbox
)]
    public class CourseCostLineSequence : PluginBase
    {
        protected override void ExecuteCDSPlugin(LocalPluginContext localcontext)
        {
            var service = localcontext.OrganizationService;
            var tracingService = localcontext.TracingService;
            tracingService.Trace("CourseCostLineSequence: Start");
            try
            {
                //Context                 
                Guid coursecostlineid = localcontext.PluginExecutionContext.PrimaryEntityId;
                var coursecostlineEntity = localcontext.Target.ToEntity<men_coursecostline>();               

                //Check Contents of Entity               
                if (!coursecostlineEntity.Contains("men_courseorderid") && !coursecostlineEntity.Contains("men_courseid"))
                {
                    tracingService.Trace("CourseCostLineSequence: No Course or Course Order in Context");
                    return;
                }

                //Check for Sequence
                if (coursecostlineEntity.Contains("men_sequence"))
                {
                    if (coursecostlineEntity.men_sequence != null)
                    {
                        tracingService.Trace("CourseCostLineSequence: Sequence={0}", coursecostlineEntity.men_sequence.Value.ToString());
                        return;
                    }
                }

                //find number of courses
                QueryExpression countcoursecostlinequery = new QueryExpression
                {
                    EntityName = men_coursecostline.EntityLogicalName,
                    ColumnSet = new ColumnSet(false)
                };
                if (coursecostlineEntity.Contains("men_courseorderid"))
                {
                    countcoursecostlinequery.Criteria.AddCondition("men_courseorderid", ConditionOperator.Equal, coursecostlineEntity.men_courseorderid.Id);
                }
                else if (coursecostlineEntity.Contains("men_courseid"))
                {
                    countcoursecostlinequery.Criteria.AddCondition("men_courseid", ConditionOperator.Equal, coursecostlineEntity.men_courseid.Id);
                }
                EntityCollection coursecostlines = service.RetrieveMultiple(countcoursecostlinequery);
                tracingService.Trace("CourseCostLineSequence: Count={0}", coursecostlines.Entities.Count.ToString());
                int sequence = coursecostlines.Entities.Count + 1;
                coursecostlineEntity["men_sequence"] = sequence;
            }            
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace("CourseCostLineSequence: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the CourseCostLineSequence plugin " + ex.Message, ex);
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace("CourseCostLineSequence: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("A standard exception error occurred in the CourseCostLineSequence plugin " + ex.Message, ex);
            }
            tracingService.Trace("CourseCostLineSequence: End");
        }
    }
}
