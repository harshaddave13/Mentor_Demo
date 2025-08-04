using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net.Http;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

using Microsoft.Crm.Sdk;
using Microsoft.Crm.Sdk.Messages;

namespace Mentor.PlugIns
{

    [CrmPluginRegistration(
        "Update",
        men_course.EntityLogicalName,
        StageEnum.PreOperation,
        ExecutionModeEnum.Synchronous,
        "men_pricebookid",
        "Course: Course Conversion Pricebook Change",
        1000,
        IsolationModeEnum.Sandbox
        , Image1Type = ImageTypeEnum.PreImage
        , Image1Name = "PreImage"
        , Image1Attributes = "men_numberofdelegates,men_pricebookid,men_conversioninfoid,men_truckcategoryid,men_numberofdelegates"
        , Description = "Course: Course Conversion Pricebook Change"
        , Id = "9F84AF80-6A90-447F-8A94-38073EE9991D"
)]
    public class CourseConversionPricebookChange : PluginBase
    {
        protected override void ExecuteCDSPlugin(LocalPluginContext localcontext)
        {
            localcontext.Trace("CourseConversionPricebookChange: Start");

            var service = localcontext.OrganizationService;
            var tracingService = localcontext.TracingService;

            var courseId = localcontext.PluginExecutionContext.PrimaryEntityId;
            var courseTarget = localcontext.Target.ToEntity<men_course>();
            var courseEntity = localcontext.MergedPreTarget.ToEntity<men_course>();

            if (!courseEntity.Contains("men_conversioninfoid") || (courseEntity.Contains("men_conversioninfoid") && courseEntity.men_conversioninfoid == null))
                return;

            var pbQuery = String.Format(@"<fetch>
                                            <entity name='men_pricebook' >
                                            <attribute name='men_days' />
                                            <attribute name='men_hours' />
                                            <attribute name='men_numberofdelegates' />
                                            <filter>
                                                <condition attribute='men_pricebookid' operator='eq' value='{0}' />
                                            </filter>
                                            <link-entity name='men_trucktype' from='men_trucktypeid' to='men_trucktypeid' alias='tt' >
                                                <attribute name='men_truckcategoryid' />
                                            </link-entity>
                                            <link-entity name='men_experiencelevel' from='men_experiencelevelid' to='men_experiencelevelid' alias='el' >
                                              <attribute name='men_name' />
                                            </link-entity>
                                            </entity>
                                        </fetch>", courseTarget.men_pricebookid.Id.ToString());
            EntityCollection pbResult = service.RetrieveMultiple(new FetchExpression(pbQuery));

            if (pbResult.Entities.Count == 0)
            {
                localcontext.Trace("CourseConversionPricebookChange: No Pricebook found - ***This should never be the case!!***");
                return;
            }

            men_pricebook pb = pbResult.Entities[0].ToEntity<men_pricebook>();

            // Get Conversion Record
            tracingService.Trace("CourseConversionPricebookChange: Get Conversion Record");
            men_conversionmatrix cm = service.Retrieve(men_conversionmatrix.EntityLogicalName, courseEntity.men_conversioninfoid.Id, new ColumnSet("men_convertfromtruckcategoryid","men_converttotruckcategoryid", "men_days", "men_hours")).ToEntity<men_conversionmatrix>();

            EntityReference tt = (EntityReference)pb.GetAttributeValue<AliasedValue>("tt.men_truckcategoryid").Value;

            if (tt == cm.men_converttotruckcategoryid && pb.men_numberofdelegates.Value == cm.men_noofdelegates.Value)
            {
                localcontext.Trace("CourseConversionPricebookChange: No Change to Conversion Info Required.");
                return;
            }
            else
            {
                if (pb.GetAttributeValue<AliasedValue>("el.men_name").Value.ToString() == "Conversion")
                {
                    localcontext.Trace("CourseConversionPricebookChange: Change to Conversion Info Required.");
                    var convMatrixQuery = String.Format(@"<fetch>
                                                      <entity name='men_conversionmatrix'>
                                                        <attribute name='men_days' />
                                                        <attribute name='men_hours' />
                                                        <filter>
                                                          <condition attribute='statecode' operator='eq' value='0' />
                                                          <condition attribute='men_convertfromtruckcategoryid' operator='eq' value='{0}' />
                                                          <condition attribute='men_converttotruckcategoryid' operator='eq' value='{1}' />
                                                          <condition attribute='men_noofdelegates' operator='eq' value='{2}' />
                                                        </filter>
                                                      </entity>
                                                    </fetch>", cm.men_convertfromtruckcategoryid.Id.ToString(), tt.Id.ToString(), pb.men_numberofdelegates.Value.ToString());
                    EntityCollection conMatrixResult = service.RetrieveMultiple(new FetchExpression(convMatrixQuery));
                    if (conMatrixResult.Entities.Count > 0)
                    {
                        localcontext.Trace("CourseConversionPricebookChange: New Conversion Info found.");
                        men_conversionmatrix conMatrix = conMatrixResult.Entities[0].ToEntity<men_conversionmatrix>();
                        courseTarget.men_days = conMatrix.men_days;
                        courseTarget.men_hours = conMatrix.men_hours;
                        courseTarget.men_conversioninfoid = new EntityReference(men_conversionmatrix.EntityLogicalName, conMatrix.Id);
                    }
                    else
                    {
                        localcontext.Trace("CourseConversionPricebookChange: No Conversion Info found, clearing value.");
                        courseTarget.men_conversioninfoid = null;
                    }
                }
                else
                {
                    localcontext.Trace("CourseConversionPricebookChange: Conversion Info needs clearing.");
                    courseTarget.men_conversioninfoid = null;
                    courseTarget.men_days = pb.men_days;
                    courseTarget.men_hours = pb.men_hours;
                }

                tracingService.Trace("CourseConversionPricebookChange: End");
            }
        }
    }
}
