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
    [CrmPluginRegistration(
    "Update",
    men_courseorder.EntityLogicalName,
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "men_bookingcount",
    "Mentor.PlugIns.CourseMapping: Update of men_courseorder",
    1000,
    IsolationModeEnum.Sandbox
    , Image1Type = ImageTypeEnum.PostImage
    , Image1Name = "PostImage"
    , Image1Attributes = "men_billingmethod, men_bookingcount, men_courseorderid, men_bookedenddate, men_ponumber, men_purchaseorderid, men_shiftid, men_bookedstartdate, statuscode, men_days, men_traveldayafter, men_traveldaybefore"
)]
    public class CourseMapping : PluginBase
    {
        protected override void ExecuteCDSPlugin(LocalPluginContext localcontext)
        {            
            var service = localcontext.OrganizationService;
            var tracingService = localcontext.TracingService;
            tracingService.Trace("CourseMapping: Start");

            try
            {
                Guid courseOrderid = localcontext.PluginExecutionContext.PrimaryEntityId;
                var courseOrderEntity = localcontext.Target.ToEntity<men_courseorder>();
                var postCourseOrder = localcontext.MergedPostTarget.ToEntity<men_courseorder>();                

                // Check if the Course Order Status is Cancelled Fully or Non chargeable or Complete
                tracingService.Trace("CourseMapping: Checking if Status if Cancelled or Completed.");
                if (postCourseOrder.statuscode.Value == men_courseorder_statuscode.CancelledFullyChargeable || postCourseOrder.statuscode.Value == men_courseorder_statuscode.Completed)
                {
                    tracingService.Trace("CourseMapping: Course Order is Cancelled or Completed.");
                    return;
                }

                if (postCourseOrder.Contains("men_bookingcount"))
                {
                    // Check that the booking count is equal to the number of days
                    tracingService.Trace("CourseMapping: Checking if Booking Count is equal to number of days.");
                    if ((int)postCourseOrder["men_bookingcount"] < decimal.Round(postCourseOrder.men_days.Value))
                    {
                        tracingService.Trace("CourseMapping: Number of bookings don't equal the number of days for the course order.");
                        return;
                    }
                }
                else
                {
                    tracingService.Trace("CourseMapping: localcontext/ Image does not include Booking Count.");
                    return;
                }
                
                // Get local Time zone - THIS WAS FOR TESTING
                //TimeZone localZone = TimeZone.CurrentTimeZone;
                //int currentYear = DateTime.Now.Year;
                //tracingService.Trace("CourseMapping: Standard Time Name: {0}, Daylight saving time name: {1}", localZone.StandardName, localZone.DaylightName);

                tracingService.Trace("CourseMapping: Creating CO to update");
                Entity updCourseOrder = new Entity(men_courseorder.EntityLogicalName);
                updCourseOrder["men_courseorderid"] = localcontext.PluginExecutionContext.PrimaryEntityId;
                updCourseOrder["men_hardbookingtimestamp"] = DateTime.Now;

                
                tracingService.Trace("CourseMapping: Bookings for Course order");

                string bookingsFetchXML = String.Format(@"<fetch>
                                                          <entity name='bookableresourcebooking'>
                                                            <attribute name='starttime' />
                                                            <attribute name='endtime' />
                                                            <attribute name='name' />
                                                            <attribute name='resource' />
                                                            <attribute name='men_instructorbookingid' />
                                                            <attribute name='duration' />
                                                            <filter>
                                                              <condition attribute='statecode' operator='eq' value='0' />
                                                              <condition attribute='bookingstatus' operator='ne' value='{0}' />
                                                              <condition attribute='men_courseorderid' operator='eq' value='{1}' />
                                                            </filter>
                                                            <order attribute='starttime' />
                                                            <link-entity name='bookableresource' from='bookableresourceid' to='resource' link-type='inner' alias='BRBbr'>
                                                              <attribute name='name' />
                                                            </link-entity>
                                                          </entity>
                                                        </fetch>", "0adbf4e6-86cc-4db0-9dbb-51b7d1ed4020", courseOrderEntity.Id.ToString());

                EntityCollection bookings = service.RetrieveMultiple(new FetchExpression(bookingsFetchXML));

                if (bookings.Entities.Count == 0)
                {
                    tracingService.Trace("CourseMapping: No Bookings for Course order");
                    return;
                }
                
                               
                tracingService.Trace("CourseMapping: Courses for Course order");
                ColumnSet coursecols = new ColumnSet(false);
                coursecols.AddColumn("men_sequence");
                coursecols.AddColumn("men_hours");
                coursecols.AddColumn("men_name");
                coursecols.AddColumn("men_courseorderid");

                QueryExpression coursequery = new QueryExpression
                {
                    EntityName = men_course.EntityLogicalName,
                    ColumnSet = coursecols
                };

                coursequery.Criteria.AddFilter(LogicalOperator.And);
                coursequery.Criteria.AddCondition("statecode", ConditionOperator.Equal, (int)men_courseState.Active);
                coursequery.Criteria.AddCondition("men_courseorderid", ConditionOperator.Equal, localcontext.PluginExecutionContext.PrimaryEntityId);

                coursequery.AddOrder("men_sequence", OrderType.Ascending);

                EntityCollection courses = service.RetrieveMultiple(coursequery);
                tracingService.Trace("CourseMapping: Courses Retreived {0}", courses.Entities.Count);

                if (courses.Entities.Count == 0)
                {
                    tracingService.Trace("CourseMapping: No Courses for Course order");
                    return;
                }
                
                                
                // Check if there is a travel day after
                bool traveldayafter = false;
                if (postCourseOrder.Contains("men_traveldayafter"))
                    traveldayafter = ((bool)postCourseOrder["men_traveldayafter"]);

                // Check if there is a travel day before
                bool traveldaybefore = false;
                if (postCourseOrder.Contains("men_traveldaybefore"))
                    traveldaybefore = ((bool)postCourseOrder["men_traveldaybefore"]);

                tracingService.Trace("CourseMapping: Adding Bookings to hashtable.");
                Hashtable bookingmappingStatus = new Hashtable();
                int bCount = 1;
                int travelDayCount = 0;
                foreach (Entity booking in bookings.Entities)
                {
                    tracingService.Trace("CourseMapping: Adding Booking Id : " + booking.Id.ToString() + " to Hashtable");
                    if ((bCount == 1 && traveldaybefore) || (bCount == bookings.Entities.Count && traveldayafter))
                    {
                        bookingmappingStatus.Add(booking.Id, true);
                        travelDayCount++;
                    }
                    else
                        bookingmappingStatus.Add(booking.Id, false);
                    bCount++;
                }
                tracingService.Trace("CourseMapping: Number of Bookings in hashtable: {0}", bookingmappingStatus.Count.ToString());


                // 2020-11-28 - AS - No longer needed as using booking duration instead of shift
                // If shift is set then get it to work out hours
                //msdyn_timegroup fulfillmentPref = null;
                decimal dailyHours = 8.00M;
                int starttime = 480;
                //if (postCourseOrder.men_shiftid != null)
                //{
                //    tracingService.Trace("CourseMapping: Shift for Course order = {0}", postCourseOrder.men_shiftid.Id.ToString());

                //    ColumnSet shiftcols = new ColumnSet(false);
                //    shiftcols.AddColumn("men_timegroupid");
                //    shiftcols.AddColumn("men_starttime");

                //    Entity shift = service.Retrieve(men_shift.EntityLogicalName, postCourseOrder.men_shiftid.Id, shiftcols);

                //    // Need to get start time from the shift
                //    if (shift.Contains("men_starttime"))
                //        starttime = ((OptionSetValue)shift["men_starttime"]).Value;

                //    ColumnSet fulfillmentCols = new ColumnSet(false);
                //    fulfillmentCols.AddColumn("msdyn_interval");

                //    if (shift.Contains("men_timegroupid"))
                //    {
                //        fulfillmentPref = (msdyn_timegroup)service.Retrieve(msdyn_timegroup.EntityLogicalName, ((EntityReference)shift["men_timegroupid"]).Id, fulfillmentCols);
                //        if (fulfillmentPref != null)
                //        {
                //            if (fulfillmentPref.msdyn_interval.HasValue)
                //            {
                //                if (fulfillmentPref.msdyn_interval < 1440)
                //                    dailyHours = fulfillmentPref.msdyn_interval.Value / 60;
                //            }
                //        }
                //    }
                //}
               

                tracingService.Trace("CourseMapping: Daily Hours = {0}", dailyHours.ToString());

                // Loop through all Bookings for Course Order with a nested loop of courses and link course to booking
                decimal currCourseHoursRemaining = 0.00M;
                decimal currDailyHoursRemaining = dailyHours;
                decimal courseHours;
                int courseCount = 0;
                int bookingCount = 1 + travelDayCount;
                int courseDayNumber = 1;
                decimal dailyHoursUsed = 0.00M;
                bool startofday = true;

                // Delete all course bookings if already exist before re-applying them
                DeleteCourseBookingsforCourseOrder(service, tracingService, localcontext.PluginExecutionContext.PrimaryEntityId);

                TimeZoneInfo utc = TimeZoneInfo.FindSystemTimeZoneById("UTC");
                TimeZoneInfo gmt = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");

                DateTime startBookingTime = DateTime.MinValue;

                // Variables used for getting dates for Courses and Course Order
                int counter = 1;
                DateTime lastDate = new DateTime();
                string courseDates = "";
                men_course courseToUpdate = new men_course();
                EntityCollection courseBookings = new EntityCollection();

                tracingService.Trace("CourseMapping: Retrieve unassigned Instructor Booking Team for Instructor Bookings");
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

                tracingService.Trace("CourseMapping: Retrieve unassigned Instructor Booking Team count : " + teamRetrieved.Entities.Count);
                if (teamRetrieved.Entities.Count == 0)
                {
                    tracingService.Trace("CourseMapping: Cannot find Unassigned Instructor Booking Team for Instructor Booking Creation");
                    throw new InvalidPluginExecutionException("CourseMapping: Cannot find Unassigned Instructor Booking Team for Instructor Booking Creation");
                }
                unassignedTeam = teamRetrieved.Entities[0].Id;

                Entity coursebooking = new Entity();
                tracingService.Trace("CourseMapping:  unassigned Instructor Booking Team Loop start : ");
                foreach (Entity course in courses.Entities)
                {
                    courseBookings.Entities.Clear();
                    courseToUpdate = new men_course();
                    courseToUpdate.Id = course.Id;

                    tracingService.Trace("check course[men_hours] : ");
                    courseHours = (decimal)course["men_hours"];
                    tracingService.Trace("check course[men_hours] : " + courseHours);

                    currCourseHoursRemaining = courseHours;

                    tracingService.Trace("check course[men_sequence] : ");
                    courseCount = (int)course["men_sequence"];
                    tracingService.Trace("check course[men_sequence] : " + course["men_sequence"]);

                    foreach (Entity booking in bookings.Entities)
                    {
                        if ((bool)bookingmappingStatus[booking.Id] == true)
                            continue;

                        // Some date checks for testing purposes can be deleted after
                        //DateTime testDate = (DateTime)booking["starttime"];

                        //var getSystemTimeZones = TimeZoneInfo.GetSystemTimeZones();
                        //bool isDaylight = TimeZoneInfo.Local.IsDaylightSavingTime(testDate);
                        //bool isTodayDyLight = TimeZoneInfo.Local.IsDaylightSavingTime(DateTime.Now);
                        //DateTime converttestDate = TimeZoneInfo.ConvertTime(testDate, utc, gmt);
                        //DateTime convertTodayDate = TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(DateTime.Now,DateTimeKind.Local), utc, gmt);

                        DateTime convertBookingDate = DateTime.MinValue;

                        if (startBookingTime == DateTime.MinValue)
                        {
                            //DateTime convertBookingDate = (DateTime)booking["starttime"];
                            convertBookingDate = (DateTime)booking["starttime"];
                            tracingService.Trace("CourseMapping: System Time: {0}", convertBookingDate);

                            convertBookingDate = TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(convertBookingDate, DateTimeKind.Utc), utc, gmt);
                            tracingService.Trace("CourseMapping: Converted Time: {0}", convertBookingDate);
                        }
                        else
                        {
                            convertBookingDate = startBookingTime;
                        }

                        // AS 20201125 - Fix use booking time and duration instead of shift 
                        if (startofday)
                        {
                            currDailyHoursRemaining = Decimal.Parse(((int)booking["duration"]).ToString()) / 60;
                            dailyHours = currDailyHoursRemaining;
                            starttime = convertBookingDate.Hour * 60 + convertBookingDate.Minute;
                            //starttime = (((DateTime)booking["starttime"]).Hour * 60) + (((DateTime)booking["starttime"]).Minute);

                            tracingService.Trace("CourseMapping: Start Time {0}", starttime.ToString());
                        }
                        tracingService.Trace("CourseMapping: Booking Number {0}", bookingCount.ToString());
                        tracingService.Trace("CourseMapping: Course Hours Populated.");

                        if (currCourseHoursRemaining > currDailyHoursRemaining && bookingCount < bookings.Entities.Count)
                        {
                            int endtime = starttime + Decimal.ToInt32((currDailyHoursRemaining * 60));
                            currCourseHoursRemaining -= (endtime - starttime) / 60;

                            if ((string)booking.GetAttributeValue<AliasedValue>("BRBbr.name").Value != "Course Cancelled")
                            {
                                coursebooking = CreateCourseBooking(service, tracingService, localcontext.PluginExecutionContext.PrimaryEntityId, booking, course, unassignedTeam, courseDayNumber, convertBookingDate, starttime, endtime, ref bookings);
                                courseBookings.Entities.Add(coursebooking);
                            }
                            else
                            {
                                tracingService.Trace("CourseMapping: Skip Generate of Course Booking as it's for a Course Cancelled Resource");
                            }

                            tracingService.Trace("CourseMapping: Course Hours Remaining (If Statement 1): {0}", currCourseHoursRemaining.ToString());

                            courseDayNumber++;
                            bookingCount++;
                            bookingmappingStatus[booking.Id] = true;
                            currDailyHoursRemaining = dailyHours;
                            startofday = true;
                            startBookingTime = DateTime.MinValue;
                            continue;
                        }
                        else if ((currCourseHoursRemaining <= currDailyHoursRemaining && currCourseHoursRemaining > 0) || bookingCount == bookings.Entities.Count)
                        {
                            int courseStartTime = starttime;
                            int endtime = 0;
                            //if (currDailyHoursRemaining < dailyHours)
                            //{
                            //    courseStartTime = starttime + Convert.ToInt32(dailyHoursUsed * 60);
                            //}
                            //else
                            //{
                            courseStartTime = starttime;
                            //}
                            endtime = courseStartTime + Convert.ToInt32(currCourseHoursRemaining * 60);
                            currDailyHoursRemaining -= currCourseHoursRemaining;
                            //currCourseHoursRemaining -= dailyHours;

                            tracingService.Trace("CourseMapping: Course Start Time {0}", courseStartTime.ToString());
                            if ((string)booking.GetAttributeValue<AliasedValue>("BRBbr.name").Value != "Course Cancelled")
                            {
                                coursebooking = CreateCourseBooking(service, tracingService, localcontext.PluginExecutionContext.PrimaryEntityId, booking, course, unassignedTeam, courseDayNumber, convertBookingDate, courseStartTime, endtime, ref bookings);
                                courseBookings.Entities.Add(coursebooking);
                            }
                            else
                            {
                                tracingService.Trace("CourseMapping: Skip Generate of Course Booking as it's for a Course Cancelled Resource");
                            }

                            //tracingService.Trace("CourseMapping: Course Hours Remaining (Else If Statement 1): {0}", currCourseHoursRemaining.ToString());

                            if (currDailyHoursRemaining <= 0 && bookingCount < bookings.Entities.Count)
                            {
                                bookingCount++;
                                bookingmappingStatus[booking.Id] = true;
                                currDailyHoursRemaining = dailyHours;
                                startofday = true;
                                dailyHoursUsed = 0;
                                courseDayNumber++;
                                startBookingTime = DateTime.MinValue;
                            }
                            else
                            {
                                if (courseHours < dailyHours)
                                    dailyHoursUsed += courseHours;
                                else
                                    dailyHoursUsed += currCourseHoursRemaining;
                                //currDailyHoursRemaining -= currCourseHoursRemaining;
                                currCourseHoursRemaining = 0;
                                startofday = false;
                                startBookingTime = convertBookingDate.AddMinutes(endtime - courseStartTime);
                                starttime = startBookingTime.Hour * 60 + startBookingTime.Minute;
                            }
                            //currCourseHoursRemaining -= dailyHours;
                            currCourseHoursRemaining -= (endtime - courseStartTime) / 60;
                            tracingService.Trace("CourseMapping: Course Hours Remaining (Else If Statement 2): {0}", currCourseHoursRemaining.ToString());
                            break;
                        }
                    }
                    //Loop for creating Course Dates string and Start and End Dates for Course
                    if (courseBookings.Entities.Count != 0)
                    {
                        foreach (Entity courseBooking in courseBookings.Entities)
                        {
                            // Sets the first date
                            if (counter == 1)
                            {
                                courseDates = ((DateTime)courseBooking["men_actualstarttime"]).ToString("dd/MM/yyyy") + ", ";
                                courseToUpdate.men_startdate = (DateTime)courseBooking["men_actualstarttime"];
                            }
                            else if ((DateTime)courseBooking["men_actualstarttime"] != lastDate)
                                courseDates = courseDates + ((DateTime)courseBooking["men_actualstarttime"]).ToString("dd/MM/yyyy") + ", ";

                            if (counter == courseBookings.Entities.Count)
                                courseToUpdate.men_enddate = (DateTime)courseBooking["men_actualendtime"];
                            counter++;
                            lastDate = (DateTime)courseBooking["men_actualstarttime"];
                        }
                        // remove the ", " from the last dateTime
                        courseDates = courseDates.Remove(courseDates.Length - 2);
                        tracingService.Trace("CourseMapping: Course Start Date: {0}, Course End Date: {1}", courseToUpdate.men_startdate, courseToUpdate.men_enddate);
                        courseToUpdate.men_coursedatesoverride = courseDates;
                        service.Update(courseToUpdate);
                        tracingService.Trace("CourseMapping: Course Dates string update: {0}", courseDates);
                        counter = 1;
                        lastDate = new DateTime();
                        courseDates = "";
                    }
                }

                // Reset Variables for Course Order Dates
                counter = 1;
                lastDate = new DateTime();
                courseDates = "";
                DateTime bookedStartDate = new DateTime();
                DateTime bookedEndDate = new DateTime();
                //Loop for creating Course Dates string for Course Order
                foreach (Entity booking in bookings.Entities)
                {
                    // Sets the first date
                    if (counter == 1)
                    {
                        courseDates = ((DateTime)booking["starttime"]).ToString("dd/MM/yyyy") + ", ";
                        bookedStartDate = (DateTime)booking["starttime"];
                    }
                    else if ((DateTime)booking["starttime"] != lastDate)
                        courseDates = courseDates + ((DateTime)booking["starttime"]).ToString("dd/MM/yyyy") + ", ";

                    if (counter == bookings.Entities.Count)
                        bookedEndDate = (DateTime)booking["endtime"];
                    counter++;
                    lastDate = (DateTime)booking["starttime"];
                }
                // remove the ", " from the last dateTime
                courseDates = courseDates.Remove(courseDates.Length - 2);
                tracingService.Trace("CourseMapping: Course Order Course Dates {0}", courseDates);
                tracingService.Trace("CourseMapping: Course Order Booked Start Date {0}", bookedStartDate);
                tracingService.Trace("CourseMapping: Course Order Booked End Date {0}", bookedEndDate);

                men_courseorder updateCourseOrder = new men_courseorder() { Id = postCourseOrder.Id };
                // If it isn't already Hard booking or Cancelled Fully or Non Chargeable
                if (postCourseOrder.statuscode.Value != men_courseorder_statuscode.HardBooking && postCourseOrder.statuscode.Value != men_courseorder_statuscode.CancelledNonChargeable && postCourseOrder.statuscode.Value != men_courseorder_statuscode.CancelledFullyChargeable)
                {
                    // If Billing method == Certificate Auto Release, Invoice Post Course and Proforma and PO input then set to HB.
                    if (postCourseOrder.men_billingmethod.Value == men_mentorfinancestatus.CertificateAutoRelease || postCourseOrder.men_billingmethod.Value == men_mentorfinancestatus.ProForma || postCourseOrder.men_billingmethod.Value == men_mentorfinancestatus.InvoicePostCourse)
                    {
                        if (postCourseOrder.Contains("men_ponumber") && postCourseOrder.Contains("men_purchaseorderid"))
                        {
                            // Set to Hard Booking
                            updateCourseOrder.statuscode = men_courseorder_statuscode.HardBooking;
                            tracingService.Trace("CourseMapping: Set Course Order to Hard Booking");
                        }
                        else if (postCourseOrder.statuscode.Value != men_courseorder_statuscode.SoftBooking)
                        {
                            // Set to Soft Booking
                            updateCourseOrder.statuscode = men_courseorder_statuscode.SoftBooking;
                            tracingService.Trace("CourseMapping: Set Course Order to Soft Booking");
                        }
                    }
                    else if (postCourseOrder.statuscode.Value != men_courseorder_statuscode.SoftBooking)
                    {
                        // Set to Soft Booking
                        updateCourseOrder.statuscode = men_courseorder_statuscode.SoftBooking;
                        tracingService.Trace("CourseMapping: Set Course Order to Soft Booking");
                    }
                }

                // Update Course Order dates if changed or null.
                if (!postCourseOrder.Contains("men_coursedates") || postCourseOrder.men_coursedates != courseDates)
                    updateCourseOrder.men_coursedates = courseDates;
                if (!postCourseOrder.Contains("men_bookedstartdate") || postCourseOrder.men_bookedstartdate != bookedStartDate)
                    updateCourseOrder.men_bookedstartdate = bookedStartDate;
                if (!postCourseOrder.Contains("men_bookedenddate") || postCourseOrder.men_bookedenddate != bookedEndDate)
                    updateCourseOrder.men_bookedenddate = bookedEndDate;

                // Set flags of Course Order for Course Booking Generation
                updateCourseOrder.men_coursebookingslastgeneratedtimestamp = DateTime.UtcNow;
                updateCourseOrder.men_coursebookingsrequireregeneration = false;
                service.Update(updateCourseOrder);
                tracingService.Trace("CourseMapping: Course Order updated");
            }            
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace("CourseMapping: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the CourseMapping plugin " + ex.Message, ex);
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace("CourseMapping: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("A standard exception error occurred in the CourseMapping plugin " + ex.Message, ex);
            }            

            tracingService.Trace("CourseMapping: End");
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
        private Entity CreateCourseBooking(IOrganizationService service, ITracingService tracingService, Guid coId, Entity booking, Entity course, Guid unassignedTeam, int dayNumber, DateTime startDate, int starttime, int endtime, ref EntityCollection bookings)
        {
            //DateTime fullStartTime = DateTime.Parse(((DateTime)booking["starttime"]).ToString("yyyy-MM-dd")).AddMinutes(starttime);
            //DateTime fullEndTime = DateTime.Parse(((DateTime)booking["starttime"]).ToString("yyyy-MM-dd")).AddMinutes(endtime);
            DateTime fullStartTime = DateTime.Parse(startDate.ToString("yyyy-MM-dd")).AddMinutes(starttime);
            DateTime fullEndTime = DateTime.Parse(startDate.ToString("yyyy-MM-dd")).AddMinutes(endtime);
            tracingService.Trace("CourseMapping: Day Number: {0}, Course: {1}, Start Time: {2}", dayNumber.ToString(), course["men_name"], fullStartTime.ToString());
            // Find Instructor
            Guid instructorid = FindInstructorForBooking(service, booking);
            tracingService.Trace("CourseMapping: Instructor = {0}", instructorid.ToString());
            //tracingService.Trace("CourseMapping: Check if a course Booking already exists.");
            //ColumnSet coursebookingcols = new ColumnSet(false);
            //coursebookingcols.AddColumn("men_bookingid");
            //coursebookingcols.AddColumn("men_courseid");
            //coursebookingcols.AddColumn("men_name");
            //coursebookingcols.AddColumn("men_courseorderid");
            //QueryExpression coursebookingquery = new QueryExpression
            //{
            //    EntityName = "men_coursebooking",
            //    ColumnSet = coursebookingcols
            //};
            //coursebookingquery.Criteria.AddFilter(LogicalOperator.And);
            //coursebookingquery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            //coursebookingquery.Criteria.AddCondition("men_courseorderid", ConditionOperator.Equal, coId);
            //coursebookingquery.Criteria.AddCondition("men_courseid", ConditionOperator.Equal, course.Id);
            //coursebookingquery.Criteria.AddCondition("men_daynumber", ConditionOperator.Equal, dayNumber);
            //EntityCollection coursebookings = service.RetrieveMultiple(coursebookingquery);

            Entity coursebooking = new Entity("men_coursebooking");
            coursebooking["men_bookingid"] = new EntityReference(BookableResourceBooking.EntityLogicalName, booking.Id);
            coursebooking["men_courseid"] = new EntityReference(men_course.EntityLogicalName, course.Id);
            coursebooking["men_courseorderid"] = (EntityReference)course["men_courseorderid"];
            coursebooking["men_daynumber"] = dayNumber;
            coursebooking["men_actualstarttime"] = fullStartTime;
            coursebooking["men_actualendtime"] = fullEndTime;
            coursebooking["men_name"] = booking["name"] + " | " + course["men_name"];
            coursebooking["men_coursesequence"] = (int)course["men_sequence"];

            if (instructorid != Guid.Empty)
                coursebooking["men_instructorid"] = new EntityReference(men_instructor.EntityLogicalName, instructorid);

            tracingService.Trace("CourseMapping: Check if BRB has Instructor Booking Linked");
            if (!booking.Contains("men_instructorbookingid"))
            {
                tracingService.Trace("CourseMapping: Instructor Booking not populated on Booking");
                // This will also link the Instructor Booking to the BRB when created
                Guid insBookingID = CreateInstructorBooking(service, tracingService, booking, unassignedTeam, instructorid, (DateTime)booking["starttime"], (DateTime)booking["endtime"]);
                coursebooking["men_instructorbookingid"] = new EntityReference("men_instructorbooking", insBookingID);
                tracingService.Trace("CourseMapping: Instructor Booking Created");

                int index = bookings.Entities.IndexOf(booking);
                if (index != -1)
                {
                    bookings.Entities[index]["men_instructorbookingid"] = new EntityReference("men_instructorbooking", insBookingID);
                    tracingService.Trace("CourseMapping: Booking in collection updated with Instructor Booking");
                }
            }
            else
            {
                tracingService.Trace("CourseMapping: Instructor Booking already populated");
                coursebooking["men_instructorbookingid"] = booking["men_instructorbookingid"];
            }

            Guid coursebookingid;
            //if (coursebookings.Entities.Count > 0)
            //{
            //    coursebooking.Id = coursebookings.Entities[0].Id;
            //    service.Update(coursebooking);.
            //    coursebookingid = coursebooking.Id;
            //    tracingService.Trace("CourseMapping: Course Booking Updated: {0}", coursebookingid.ToString());
            //}
            //else
            //{
            coursebookingid = service.Create(coursebooking);
            tracingService.Trace("CourseMapping: Course Booking Created: {0}", coursebookingid.ToString());
            //}
            return coursebooking;
        }
        private Guid CreateInstructorBooking(IOrganizationService service, ITracingService tracingService, Entity booking, Guid team, Guid instructorid, DateTime fullStartTime, DateTime fullEndTime)
        {
            men_instructorbooking instructorBooking = new men_instructorbooking();

            instructorBooking.men_originalstart = fullStartTime;
            instructorBooking.men_originalend = fullEndTime;
            instructorBooking.men_bookingid = new EntityReference("bookableresourcebooking", booking.Id);
            instructorBooking.men_bookingtype = men_mentorbookingtypes.SiteBooking;
            instructorBooking.men_name = "Course";
            if (instructorid != Guid.Empty)
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
            instructorBooking.OwnerId = new EntityReference("team", team);

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
        /// <summary>
        /// Delete all previous course bookings before applying new ones.
        /// </summary>
        /// <param name="service"></param>
        /// <param name="tracingService"></param>
        /// <param name="coId"></param>
        private void DeleteCourseBookingsforCourseOrder(IOrganizationService service, ITracingService tracingService, Guid coId)
        {
            ColumnSet coursebookingcols = new ColumnSet(false);
            coursebookingcols.AddColumn("men_bookingid");
            coursebookingcols.AddColumn("men_courseid");
            coursebookingcols.AddColumn("men_name");
            coursebookingcols.AddColumn("men_courseorderid");

            QueryExpression coursebookingquery = new QueryExpression
            {
                EntityName = "men_coursebooking",
                ColumnSet = coursebookingcols
            };

            coursebookingquery.Criteria.AddFilter(LogicalOperator.And);
            coursebookingquery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            coursebookingquery.Criteria.AddCondition("men_courseorderid", ConditionOperator.Equal, coId);

            EntityCollection coursebookings = service.RetrieveMultiple(coursebookingquery);

            if (coursebookings.Entities.Count > 0)
            {
                tracingService.Trace("CourseMapping: Deleting previous Course Bookings");

                foreach (Entity ent in coursebookings.Entities)
                {
                    service.Delete("men_coursebooking", ent.Id);
                }
            }
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
        private Guid GetInstructoBookingFromBRB(IOrganizationService service, ITracingService tracingService, Entity booking)
        {
            Guid instructorBooking = Guid.Empty;
            string brbFetchXML = String.Format(@"<fetch top='1'>
                                                  <entity name='bookableresourcebooking'>
                                                    <attribute name='bookableresourcebookingid' />
                                                    <attribute name='men_instructorbookingid' />
                                                    <filter>
                                                      <condition attribute='bookableresourcebookingid' operator='eq' value='{0}' />
                                                      <condition attribute='men_instructorbookingid' operator='not-null' />
                                                    </filter>
                                                  </entity>
                                                </fetch>", booking.Id.ToString());
            EntityCollection brbRetrieved = service.RetrieveMultiple(new FetchExpression(brbFetchXML));

            if (brbRetrieved.Entities.Count > 0)
                instructorBooking = brbRetrieved.Entities[0].GetAttributeValue<EntityReference>("men_instructorbookingid").Id;

            return instructorBooking;
        }
    }
}
