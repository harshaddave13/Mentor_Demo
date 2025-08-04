using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace Mentor.PlugIns
{
    [CrmPluginRegistration("Update",
men_courseorder.EntityLogicalName, StageEnum.PostOperation, ExecutionModeEnum.Synchronous,
"statuscode", "Mentor.PlugIns.TrainingCentreCourseMapping: Update of men_courseorder", 1000,
IsolationModeEnum.Sandbox
)]

    [CrmPluginRegistration("Update",
men_trainingcentrecourse.EntityLogicalName, StageEnum.PostOperation, ExecutionModeEnum.Synchronous,
"statuscode", "Mentor.PlugIns.TrainingCentreCourseMapping: Update of men_trainingcentrecourse", 1000,
IsolationModeEnum.Sandbox
    , Image1Type = ImageTypeEnum.PostImage
, Image1Name = "PostImage"
, Image1Attributes = "men_days, men_numberofdaysbooked, statuscode, men_trainingcentrecourseid"
)]
    public class TrainingCentreCourseMapping : PluginBase
    {
        protected override void ExecuteCDSPlugin(LocalPluginContext localcontext)
        {            
            var service = localcontext.OrganizationService;
            var tracingService = localcontext.TracingService;
            tracingService.Trace("TrainingCentreCourseMapping: Start");
            try
            {
                Guid trainingCentreBookingid = localcontext.PluginExecutionContext.PrimaryEntityId;
                var trainingCentreBookingEntity = localcontext.Target.ToEntity<men_trainingcentrecourse>();
                var postTrainingCentreBooking = localcontext.MergedPostTarget.ToEntity<men_trainingcentrecourse>();


                if (postTrainingCentreBooking.statuscode.Value != men_trainingcentrecourse_statuscode.Booked)
                {
                    tracingService.Trace("TrainingCentreCourseMapping: Not set to booked");
                    return;
                }

                if (postTrainingCentreBooking.men_numberofdaysbooked.Value < postTrainingCentreBooking.men_days.Value)
                {
                    tracingService.Trace("TrainingCentreCourseMapping: Not all Training Centre Booking days are booked");
                    throw new InvalidPluginExecutionException("Not all Training Centre Booking days are booked");
                }

                // First get training centre courses
                string fetchTrainingCentreCourseDetails = String.Format(@"<fetch>
                                                                              <entity name='men_trainingcentrecoursedetail' >
                                                                                <attribute name='men_sequence' />
                                                                                <attribute name='men_hours' />
                                                                                <attribute name='men_name' />
                                                                                <attribute name='men_trainingcentrecoursedetailid' />
                                                                                <filter>
                                                                                  <condition attribute='statecode' operator='eq' value='0' />
                                                                                  <condition attribute='men_trainingcentrecourse' operator='eq' value='{0}' />
                                                                                </filter>
                                                                                <order attribute='men_sequence' />
                                                                              </entity>
                                                                            </fetch>", postTrainingCentreBooking.Id.ToString());

                EntityCollection retrievedTrainingCentreCourseDetails = service.RetrieveMultiple(new FetchExpression(fetchTrainingCentreCourseDetails));

                if (retrievedTrainingCentreCourseDetails.Entities.Count == 0)
                {
                    tracingService.Trace("TrainingCentreCourseMapping: No Training Centre Courses on Training Centre Booking");
                    return;
                }

                //Get Bookable Resource Bookings
                string fetchBookings = String.Format(@"<fetch>
                                                          <entity name='bookableresourcebooking' >
                                                            <attribute name='bookableresourcebookingid' />
                                                            <attribute name='starttime' />
                                                            <attribute name='endtime' />
                                                            <attribute name='name' />
                                                            <attribute name='resource' />
                                                            <attribute name='men_instructorbookingid' />
                                                            <filter type='and' >
                                                              <condition attribute='men_trainingcentrecourseid' operator='eq' value='{0}' />
                                                              <condition attribute='statecode' operator='eq' value='0' />
                                                              <condition attribute='men_instructorid' operator='not-null' />
                                                            </filter>
                                                            <order attribute='starttime' />
                                                            <link-entity name='bookingstatus' from='bookingstatusid' to='bookingstatus' link-type='inner' >
                                                              <filter>
                                                                <condition attribute='name' operator='eq' value='Hard Booking' />
                                                              </filter>
                                                            </link-entity>
                                                          </entity>
                                                        </fetch>", postTrainingCentreBooking.Id.ToString());

                EntityCollection retrievedBookings = service.RetrieveMultiple(new FetchExpression(fetchBookings));

                if (retrievedBookings.Entities.Count == 0)
                {
                    tracingService.Trace("TrainingCentreCourseMapping: No Bookings for Training Centre Booking");
                    return;
                }

                tracingService.Trace("TrainingCentreCourseMapping: Adding Bookings to hashtable.");
                Hashtable bookingmappingStatus = new Hashtable();
                foreach (Entity booking in retrievedBookings.Entities)
                {
                    tracingService.Trace("TrainingCentreCourseMapping: Adding Booking Id : {0} to Hashtable", booking.Id.ToString());
                    bookingmappingStatus.Add(booking.Id, false);
                }
                tracingService.Trace("TrainingCentreCourseMapping: Number of Bookings in hashtable: {0}", bookingmappingStatus.Count.ToString());

                //Shift
                decimal dailyHours = 8.00M;

                tracingService.Trace("TrainingCentreCourseMapping: Daily Hours = {0}", dailyHours.ToString());

                // Loop through all Bookings for Training Centre Booking with a nested loop of courses and link course to booking
                decimal currCourseHoursRemaining = 0.00M;
                decimal currDailyHoursRemaining = dailyHours;
                decimal courseHours = 0.00M;
                int bookingCount = 1;
                int courseDayNumber = 1;
                decimal dailyHoursUsed = 0;

                TimeZoneInfo utc = TimeZoneInfo.FindSystemTimeZoneById("UTC");
                TimeZoneInfo gmt = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");

                DateTime convertBookingDate = (DateTime)retrievedBookings.Entities[0]["starttime"];
                tracingService.Trace("TrainingCentreCourseMapping: System Time: {0}", convertBookingDate);

                convertBookingDate = TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(convertBookingDate, DateTimeKind.Utc), utc, gmt);
                tracingService.Trace("TrainingCentreCourseMapping: Converted Time: {0}", convertBookingDate);

                tracingService.Trace("TrainingCentreCourseMapping: Retrieve unassigned Instructor Booking Team for Instructor Bookings");
                Guid unassignedTeam = new Guid();
                string teamFetchXML = String.Format(@"<fetch>
                                                        <entity name='team'>
                                                        <attribute name='teamid' />
                                                        <filter>
                                                            <condition attribute='name' operator='eq' value='Unassigned Instructor Booking Team' />
                                                        </filter>
                                                        </entity>
                                                    </fetch>");
                EntityCollection teamRetrieved = service.RetrieveMultiple(new FetchExpression(teamFetchXML));

                if (teamRetrieved.Entities.Count == 0)
                {
                    tracingService.Trace("TrainingCentreCourseMapping: Cannot find Unassigned Instructor Booking Team for Instructor Booking Creation");
                    throw new InvalidPluginExecutionException("TrainingCentreCourseMapping: Cannot find Unassigned Instructor Booking Team for Instructor Booking Creation");
                }

                unassignedTeam = teamRetrieved.Entities[0].Id;

                foreach (Entity TCcourse in retrievedTrainingCentreCourseDetails.Entities)
                {
                    if (TCcourse.Contains("men_hours"))
                    {
                        currCourseHoursRemaining = (decimal)TCcourse["men_hours"];
                    }
                    else
                    {
                        currCourseHoursRemaining = 0;
                    }

                    int starttime = convertBookingDate.Hour * 60 + convertBookingDate.Minute;
                    Entity coursebooking = new Entity("men_coursebooking");
                    foreach (Entity booking in retrievedBookings.Entities)
                    {
                        if ((bool)bookingmappingStatus[booking.Id] == true)
                            continue;

                        tracingService.Trace("TrainingCentreCourseMapping: Booking Number {0}", bookingCount.ToString());

                        tracingService.Trace("TrainingCentreCourseMapping: Course Hours Populated.");

                        if (currCourseHoursRemaining > currDailyHoursRemaining)
                        {
                            int endtime = starttime + Decimal.ToInt32((currDailyHoursRemaining * 60));
                            currCourseHoursRemaining -= (endtime - starttime) / 60;
                            Guid coursebookingid = CreateCourseBooking(service, tracingService, localcontext.PluginExecutionContext.PrimaryEntityId, booking, unassignedTeam, TCcourse, postTrainingCentreBooking, courseDayNumber, starttime, endtime);
                            tracingService.Trace("TrainingCentreCourseMapping: Course Hours Remaining (If Statement 1): {0}", currCourseHoursRemaining.ToString());
                            courseDayNumber++;
                            bookingCount++;
                            bookingmappingStatus[booking.Id] = true;
                            currDailyHoursRemaining = dailyHours;
                            continue;
                        }
                        else if ((currCourseHoursRemaining <= currDailyHoursRemaining && currCourseHoursRemaining >= 0) || bookingCount == retrievedBookings.Entities.Count)
                        {

                            int courseStartTime = starttime;
                            int endtime = 0;
                            if (currDailyHoursRemaining < dailyHours)
                            {
                                courseStartTime = starttime + (Convert.ToInt32(dailyHoursUsed) * 60);
                            }
                            else
                            {
                                courseStartTime = starttime;
                            }

                            endtime = courseStartTime + (Convert.ToInt32(currCourseHoursRemaining) * 60);

                            currDailyHoursRemaining -= currCourseHoursRemaining;
                            currCourseHoursRemaining -= dailyHours;

                            tracingService.Trace("TrainingCentreCourseMapping: Course Start Time {0}", courseStartTime.ToString());

                            Guid coursebookingid = CreateCourseBooking(service, tracingService, localcontext.PluginExecutionContext.PrimaryEntityId, booking, unassignedTeam, TCcourse, postTrainingCentreBooking, courseDayNumber, courseStartTime, endtime);
                            tracingService.Trace("TrainingCentreCourseMapping: Course Hours Remaining (Else If Statement 1): {0}", currCourseHoursRemaining.ToString());

                            if (currDailyHoursRemaining <= 0 && bookingCount < retrievedBookings.Entities.Count)
                            {
                                bookingCount++;
                                bookingmappingStatus[booking.Id] = true;
                                currDailyHoursRemaining = dailyHours;
                                dailyHoursUsed = 0;
                                courseDayNumber++;
                            }
                            else
                            {
                                if (courseHours < dailyHours)
                                    dailyHoursUsed += courseHours;
                                else
                                    dailyHoursUsed += currCourseHoursRemaining;

                                //currDailyHoursRemaining -= currCourseHoursRemaining;
                                currCourseHoursRemaining = 0;
                            }
                            break;
                        }

                    }

                }
            }            
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace("TrainingCentreCourseMapping: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the CourseMapping plugin " + ex.Message, ex);
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace("TrainingCentreCourseMapping: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("A standard exception error occurred in the CourseMapping plugin " + ex.Message, ex);
            }
           

            tracingService.Trace("TrainingCentreCourseMapping: End");
        }

        /// <summary>
        /// Creates the Course Booking record
        /// </summary>
        /// <param name="service"></param>
        /// <param name="booking"></param>
        /// <param name="course"></param>
        /// <param name="dayNumber"></param>
        /// <param name="starttime"></param>
        /// <returns></returns>
        private Guid CreateCourseBooking(IOrganizationService service, ITracingService tracingService, Guid coId, Entity booking, Guid unassignedTeamid, Entity trainingCentreCourse, Entity trainingCentreBooking, int dayNumber, int starttime, int endtime)
        {

            DateTime fullStartTime = DateTime.Parse(((DateTime)booking["starttime"]).ToString("yyyy-MM-dd")).AddMinutes(starttime);
            DateTime fullEndTime = DateTime.Parse(((DateTime)booking["starttime"]).ToString("yyyy-MM-dd")).AddMinutes(endtime);
            tracingService.Trace("TrainingCentreCourseMapping: Day Number: {0}, Course: {1}, Start Time: {2}", dayNumber.ToString(), trainingCentreCourse["men_name"], fullStartTime.ToString());

            // Find Instructor
            Guid instructorid = FindInstructorForBooking(service, booking);
            tracingService.Trace("TrainingCentreCourseMapping: Instructor = {0}", instructorid.ToString());

            tracingService.Trace("TrainingCentreCourseMapping: Check if a course Booking already exists.");

            ColumnSet coursebookingcols = new ColumnSet(false);
            coursebookingcols.AddColumn("men_bookingid");
            coursebookingcols.AddColumn("men_trainingcentrecourseid");
            coursebookingcols.AddColumn("men_name");
            coursebookingcols.AddColumn("men_courseorderid");

            QueryExpression coursebookingquery = new QueryExpression
            {
                EntityName = "men_coursebooking",
                ColumnSet = coursebookingcols
            };

            coursebookingquery.Criteria.AddFilter(LogicalOperator.And);
            coursebookingquery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            //coursebookingquery.Criteria.AddCondition("men_courseorderid", ConditionOperator.Equal, coId);
            coursebookingquery.Criteria.AddCondition("men_trainingcentrecoursedetailid", ConditionOperator.Equal, trainingCentreCourse.Id);
            coursebookingquery.Criteria.AddCondition("men_daynumber", ConditionOperator.Equal, dayNumber);

            EntityCollection coursebookings = service.RetrieveMultiple(coursebookingquery);

            Entity coursebooking = new Entity("men_coursebooking");
            coursebooking["men_bookingid"] = new EntityReference(BookableResourceBooking.EntityLogicalName, booking.Id);
            coursebooking["men_trainingcentrecoursedetailid"] = new EntityReference(men_trainingcentrecoursedetail.EntityLogicalName, trainingCentreCourse.Id);
            coursebooking["men_trainingcentrecourseid"] = new EntityReference(men_trainingcentrecourse.EntityLogicalName, trainingCentreBooking.Id);
            //coursebooking["men_courseorderid"] = (EntityReference)course["men_courseorderid"];
            coursebooking["men_daynumber"] = dayNumber;
            coursebooking["men_actualstarttime"] = fullStartTime;
            coursebooking["men_actualendtime"] = fullEndTime;
            coursebooking["men_name"] = "Course Booking | " + trainingCentreCourse["men_name"];

            if (instructorid != Guid.Empty)
                coursebooking["men_instructorid"] = new EntityReference(men_instructor.EntityLogicalName, instructorid);

            if (!booking.Contains("men_instructorbookingid"))
            {
                // This will also link the Instructor Booking to the BRB when created
                Guid insBookingID = CreateInstructorBooking(service, tracingService, booking, unassignedTeamid, instructorid, fullStartTime, fullEndTime);
                coursebooking["men_instructorbookingid"] = new EntityReference("men_instructorbooking", insBookingID);
            }
            else
            {
                coursebooking["men_instructorbookingid"] = booking["men_instructorbookingid"];
            }

            Guid coursebookingid;
            if (coursebookings.Entities.Count > 0)
            {
                coursebooking.Id = coursebookings.Entities[0].Id;
                service.Update(coursebooking);
                coursebookingid = coursebooking.Id;
                tracingService.Trace("TrainingCentreCourseMapping: Course Booking Updated: {0}", coursebookingid.ToString());
            }
            else
            {
                coursebookingid = service.Create(coursebooking);
                tracingService.Trace("TrainingCentreCourseMapping: Course Booking Created: {0}", coursebookingid.ToString());
            }

            return coursebookingid;

        }

        private Guid CreateInstructorBooking(IOrganizationService service, ITracingService tracingService, Entity booking, Guid teamid, Guid instructorid, DateTime fullStartTime, DateTime fullEndTime)
        {
            men_instructorbooking instructorBooking = new men_instructorbooking();

            instructorBooking.men_originalstart = fullStartTime;
            instructorBooking.men_originalend = fullEndTime;
            instructorBooking.men_bookingid = new EntityReference("bookableresourcebooking", booking.Id);
            instructorBooking.men_bookingtype =men_mentorbookingtypes.SiteBooking;
            instructorBooking.men_name = "Training Centre Course";
            instructorBooking.men_Instructorid = new EntityReference(men_instructor.EntityLogicalName, instructorid);

            string teamFetchXML = String.Format(@"<fetch>
                                                        <entity name='team'>
                                                        <attribute name='teamid' />
                                                        <filter>
                                                            <condition attribute='name' operator='eq' value='Unassigned Instructor Booking Team' />
                                                        </filter>
                                                        </entity>
                                                    </fetch>");
            EntityCollection unassignedTeam = service.RetrieveMultiple(new FetchExpression(teamFetchXML));

            if (unassignedTeam.Entities.Count == 0)
            {
                tracingService.Trace("CourseMapping: Cannot find Unassigned Instructor Booking Team for Instructor Booking Creation");
                throw new InvalidPluginExecutionException("CourseMapping: Cannot find Unassigned Instructor Booking Team for Instructor Booking Creation");
            }
            instructorBooking.OwnerId = new EntityReference("team", teamid);

            Guid insBookingID = service.Create(instructorBooking);
            tracingService.Trace("CourseMapping: Instructor Booking Created: {0}", insBookingID.ToString());

            tracingService.Trace("CourseMapping: Link Instructor Booking to Bookable Resource Booking");
            BookableResourceBooking bookableResourceBooking = new BookableResourceBooking();
            bookableResourceBooking.Id = booking.Id;
            bookableResourceBooking.men_instructorbookingid = new EntityReference("men_instructorbooking", insBookingID);
            service.Update(bookableResourceBooking);
            tracingService.Trace("CourseMapping: Bookable Resource Booking Updated");

            return insBookingID;
        }

        private Guid FindInstructorForBooking(IOrganizationService service, Entity booking)
        {
            Guid resource = Guid.Empty;
            if (booking.Contains("resource"))
                resource = ((EntityReference)booking["resource"]).Id;
            else
                return resource;

            ColumnSet instructorcols = new ColumnSet(false);
            instructorcols.AddColumn("men_instructorid");

            QueryExpression instructorquery = new QueryExpression
            {
                EntityName = "men_instructor",
                ColumnSet = instructorcols
            };

            instructorquery.Criteria.AddFilter(LogicalOperator.And);
            instructorquery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            instructorquery.Criteria.AddCondition("men_bookableresourceid", ConditionOperator.Equal, resource);

            EntityCollection instructors = service.RetrieveMultiple(instructorquery);

            if (instructors.Entities.Count > 0)
                return instructors.Entities[0].Id;
            else
                return Guid.Empty;
        }
    }
}
