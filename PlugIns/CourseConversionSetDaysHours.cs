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
    "men_conversioninfoid",
    "Course: Course Conversion",
    1000,
    IsolationModeEnum.Sandbox
    , Image1Type = ImageTypeEnum.PreImage
    , Image1Name = "PreImage"
    , Image1Attributes = "men_numberofdelegates,men_pricebookid"
)]
    public class CourseConversionSetDaysHours : PluginBase
    {
        protected override void ExecuteCDSPlugin(LocalPluginContext localcontext)
        {
            localcontext.Trace("CourseConversionSetDaysHours: Start");

            var service = localcontext.OrganizationService;
            var tracingService = localcontext.TracingService;

            var courseId = localcontext.PluginExecutionContext.PrimaryEntityId;
            var courseTarget = localcontext.Target.ToEntity<men_course>();
            var courseEntity = localcontext.MergedPreTarget.ToEntity<men_course>();

            if (!courseEntity.Contains("men_conversioninfoid"))
                return;

            // If the Conversion info has been cleared then retun as the platform operation will blank it
            if (courseTarget.men_conversioninfoid == null)
            {
                tracingService.Trace("CourseConversionSetDaysHours: Conversion Info has been nulled so get Days and Hours from Pricebook");
                men_pricebook pb = service.Retrieve(men_pricebook.EntityLogicalName, courseEntity.men_pricebookid.Id, new ColumnSet("men_days", "men_hours")).ToEntity<men_pricebook>();
                courseTarget.men_days = pb.men_days;
                courseTarget.men_hours = pb.men_hours;
                return;
            }

            men_conversionmatrix conversionMatrix = service.Retrieve(men_conversionmatrix.EntityLogicalName, courseEntity.men_conversioninfoid.Id, new ColumnSet("men_days", "men_hours", "men_noofdelegates", "men_convertfromtruckcategoryid", "men_converttotruckcategoryid")).ToEntity<men_conversionmatrix>();

            // If the Course Matrix is not for the correct number of Delegates then attempt to get the correct one.
            if (courseEntity.men_numberofdelegates.Value != conversionMatrix.men_noofdelegates.Value)
            {
                tracingService.Trace("CourseConversionSetDaysHours: Incorrect Conversion Matrix attempt to find the correct one for the correct number of delegates");
                // Get the conversion matrix for the correct number of delegates
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
                                                    </fetch>", conversionMatrix.men_convertfromtruckcategoryid.Id.ToString(), conversionMatrix.men_converttotruckcategoryid.Id.ToString(), courseEntity.men_numberofdelegates.Value.ToString());
                EntityCollection conMatrixResult = service.RetrieveMultiple(new FetchExpression(convMatrixQuery));
                if (conMatrixResult.Entities.Count > 0)
                {
                    tracingService.Trace("CourseConversionSetDaysHours: Found Correct Conversion Matrix set values");
                    men_conversionmatrix conMatrix = conMatrixResult.Entities[0].ToEntity<men_conversionmatrix>();
                    courseTarget.men_days = conMatrix.men_days;
                    courseTarget.men_hours = conMatrix.men_hours;
                    courseTarget.men_conversioninfoid = new EntityReference(men_conversionmatrix.EntityLogicalName, conMatrix.Id);
                }
                else
                {
                    tracingService.Trace("CourseConversionSetDaysHours: Did not find Correct Conversion Matrix set values to original Conversion Matrix");
                    courseTarget.men_days = conversionMatrix.men_days;
                    courseTarget.men_hours = conversionMatrix.men_hours;
                }
            }
            else
            {
                tracingService.Trace("CourseConversionSetDaysHours: Correct Conversion Matrix set values");
                courseTarget.men_days = conversionMatrix.men_days;
                courseTarget.men_hours = conversionMatrix.men_hours;
            }

            tracingService.Trace("CourseConversionSetDaysHours: End");

        }
    }
}
