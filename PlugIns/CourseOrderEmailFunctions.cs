using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mentor.PlugIns
{
    public class CourseOrderEmailFunctions
    {
        public static List<string> GetAccountInfo(Guid record, List<string> items, IOrganizationService service, ITracingService tracingService)
        {
            string content = String.Empty;
            string greet = String.Empty;
            string location = String.Empty;
            string tc = String.Empty;
            string bc = String.Empty;
            string ownerid = String.Empty;
            string coNumber = string.Empty;
            string startTime = string.Empty;
            tracingService.Trace("CreateQuoteEmail: start of function 1 ");
            TimeSpan noon = new TimeSpan(12, 0, 0);
            TimeSpan now = DateTime.Now.TimeOfDay;
            TimeZoneInfo utc = TimeZoneInfo.FindSystemTimeZoneById("UTC");
            TimeZoneInfo gmt = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");
            string courseOrderInfoquery = string.Format(@"<fetch>
                                                              <entity name='men_courseorder'>
                                                                <attribute name='men_courseordernumber' />
                                                                <attribute name='men_name' />
                                                                <attribute name='men_bookedstartdate' />
                                                                <attribute name='men_accountid' />
                                                                <filter>
                                                                  <condition attribute='men_courseorderid' operator='eq' value='{0}' />
                                                                </filter>
                                                                <link-entity name='account' from='accountid' to='men_accountid' link-type='outer' alias='SiteAcc'>
                                                                  <attribute name='address1_city' />
                                                                  <attribute name='address1_line1' />
                                                                  <attribute name='address1_line2' />
                                                                  <attribute name='address1_postalcode' />
                                                                  <attribute name='address2_city' />
                                                                  <attribute name='address2_line1' />
                                                                  <attribute name='address2_line2' />
                                                                  <attribute name='address2_postalcode' />
                                                                  <attribute name='men_customerspecifictandcs' />
                                                                  <attribute name='name' />
                                                                  <attribute name='msdyn_billingaccount' />
                                                                  <link-entity name='account' from='accountid' to='msdyn_billingaccount' link-type='outer' alias='BillAcc'>
                                                                    <attribute name='address1_city' />
                                                                    <attribute name='address1_line1' />
                                                                    <attribute name='address1_line2' />
                                                                    <attribute name='address1_postalcode' />
                                                                    <attribute name='name' />
                                                                    <attribute name='men_customerspecifictandcs' />
                                                                  </link-entity>
                                                                </link-entity>
                                                                <link-entity name='contact' from='contactid' to='men_contactid' link-type='outer' alias='BC'>
                                                                  <attribute name='contactid' />
                                                                  <attribute name='firstname' />
                                                                </link-entity>
                                                                <link-entity name='systemuser' from='systemuserid' to='ownerid' link-type='outer' alias='owner'>
                                                                  <attribute name='systemuserid' />
                                                                </link-entity>
                                                              </entity>
                                                            </fetch>", record.ToString());
            EntityCollection entities = service.RetrieveMultiple(new FetchExpression(courseOrderInfoquery));
            foreach (Entity e in entities.Entities)
            {
                if (e.Contains("SiteAcc.msdyn_billingaccount"))
                {
                    if (e.Contains("BillAcc.name"))
                    {
                        tracingService.Trace("CreateQuoteEmail: Billing Account name");
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("BillAcc.name").Value.ToString()}</p";
                    }
                    if (e.Contains("BillAcc.address1_line1"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("BillAcc.address1_line1").Value.ToString()}</p>";
                    }
                    if (e.Contains("BillAcc.address1_line2"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("BillAcc.address1_line2").Value.ToString()} </p>";
                    }
                    if (e.Contains("BillAcc.address1_city"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("BillAcc.address1_city").Value.ToString()} </p>";
                    }
                    if (e.Contains("BillAcc.address1_postalcode"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("BillAcc.address1_postalcode").Value.ToString()}  </p>";
                    }
                }
                else if (e.Contains("men_accountid"))
                {
                    if (e.Contains("SiteAcc.name"))
                    {
                        tracingService.Trace("CreateQuoteEmail: Site Account name");
                        content += $"<p>{e.GetAttributeValue<AliasedValue>("SiteAcc.name").Value.ToString()}</p>";
                    }
                    if (e.Contains("SiteAcc.address2_line1"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.address2_line1").Value.ToString()}</p>";
                    }
                    if (e.Contains("SiteAcc.address2_line2"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.address2_line2").Value.ToString()} </p>";
                    }
                    if (e.Contains("SiteAcc.address2_city"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.address2_city").Value.ToString()}  </p>";
                    }
                    if (e.Contains("SiteAcc.address2_postalcode"))
                    {
                        content += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.address2_postalcode").Value.ToString()}  </p>";
                    }
                }
                coNumber = e.GetAttributeValue<string>("men_courseordernumber").ToString();

                tracingService.Trace("CourseOrderConfirmation:" + content.ToString());
                if (e.Contains("men_bookedstartdate"))
                {
                    DateTime bookedStartTime = ((DateTime)e["men_bookedstartdate"]);
                    startTime = (TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(bookedStartTime, DateTimeKind.Utc), utc, gmt)).ToString("HH:mm");
                }

                if (e.Contains("BillAcc.men_customerspecifictandcs"))
                {
                    tracingService.Trace("CreateQuoteEmail:cst&c bill acc ");
                    tc = $"<p> {e.GetAttributeValue<AliasedValue>("BillAcc.men_customerspecifictandcs").Value.ToString()}</p>";

                }
                else if (e.Contains("SiteAcc.men_customerspecifictandcs"))
                {
                    tracingService.Trace("CreateQuoteEmail:cst&c site acc ");
                    tc = $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.men_customerspecifictandcs").Value.ToString()}</p>";
                }
                tracingService.Trace("CreateQuoteEmail T&C:" + tc);
                if (e.Contains("BC.firstname"))
                {
                    if (now > noon)
                    {
                        greet = $"<p> Good Afternoon, {e.GetAttributeValue<AliasedValue>("BC.firstname").Value.ToString()}</p>";
                    }
                    else
                    {
                        greet = $"<p>Good Morning, {e.GetAttributeValue<AliasedValue>("BC.firstname").Value.ToString()} </p>";
                    }
                }
                else
                {
                    if (now > noon)
                    {
                        greet = "<p> Good Afternoon</p>";
                    }
                    else
                    {
                        greet = "<p>Good Morning </p>";
                    }
                }
                tracingService.Trace("CreateQuoteEmail: Greeting:" + greet);
                if (e.Contains("SiteAcc.name"))
                {
                    location += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.name").Value.ToString()} </p>";

                }
                if (e.Contains("SiteAcc.address1_line1"))
                {
                    location += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.address1_line1").Value.ToString()} </p>";

                }
                if (e.Contains("SiteAcc.address1_line2"))
                {
                    location += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.address1_line2").Value.ToString()} </p>";

                }
                if (e.Contains("SiteAcc.address1_city"))
                {
                    location += $"<p> {e.GetAttributeValue<AliasedValue>("SiteAcc.address1_city").Value.ToString()} </p>";
                }
                if (e.Contains("SiteAcc.address1_postalcode"))
                {
                    location += $"<p>{e.GetAttributeValue<AliasedValue>("SiteAcc.address1_postalcode").Value.ToString()} </p>";
                }
                bc = e.GetAttributeValue<AliasedValue>("BC.contactid").Value.ToString();
                tracingService.Trace("CreateQuoteEmail: ToEmailAddress:" + bc.ToString());
                ownerid = e.GetAttributeValue<AliasedValue>("owner.systemuserid").Value.ToString();
                tracingService.Trace("CreateQuoteEmail: ToEmailAddress:" + ownerid.ToString());

            }
            items.Add(content);
            items.Add(greet);
            items.Add(location);
            items.Add(tc);
            items.Add(bc);
            items.Add(ownerid);
            items.Add(coNumber);
            items.Add(startTime);
            return items;
        }
        public static string GetTrainingLocation(Guid courseID, IOrganizationService service, ITracingService tracingService)
        {
            string text = String.Empty;
            tracingService.Trace("CreateQuoteEmail: Start of traininglocation funciton ");
            string trainingLocquery = string.Format(@"<fetch>
                                              <entity name='men_course'>
                                                <filter>
                                                  <condition attribute='men_courseid' operator='eq' value='{0}' />
                                                </filter>
                                                <link-entity name='men_trainingcentrecourse' from='men_trainingcentrecourseid' to='men_trainingcentrecourseid' link-type='outer' alias='TCC'>
                                                  <link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='men_organizationunitid' link-type='outer' alias='loc'>
                                                    <attribute name='men_addressline1' />
                                                    <attribute name='men_addressline2' />
                                                    <attribute name='men_city' />
                                                    <attribute name='men_postalcode' />
                                                  </link-entity>
                                                </link-entity>
                                              </entity>
                                            </fetch>", courseID.ToString());
            EntityCollection tlEntityCollection = service.RetrieveMultiple(new FetchExpression(trainingLocquery));
            text += "<p> Mentor Training Location </p>";
            foreach (Entity course in tlEntityCollection.Entities)
            {
                tracingService.Trace("CreateQuoteEmail: Inside traininglocation function loop ");
                if (course.Contains("loc.men_addressline1"))
                {
                    text += $"<p>{course.GetAttributeValue<AliasedValue>("loc.men_addressline1").Value.ToString()}</p>";
                }
                if (course.Contains("loc.men_addressline2"))
                {
                    text += $"<p>{course.GetAttributeValue<AliasedValue>("loc.men_addressline2").Value.ToString()}</p>";
                }
                tracingService.Trace("CreateQuoteEmail:address 1  ");
                tracingService.Trace("CreateQuoteEmail: address 2 ");
                if (course.Contains("loc.men_city"))
                {
                    text += $"<p>{course.GetAttributeValue<AliasedValue>("loc.men_city").Value.ToString()}</p>";
                }
                tracingService.Trace("CreateQuoteEmail: City ");
                if (course.Contains("loc.men_postalcode"))
                {
                    text += $"<p>{course.GetAttributeValue<AliasedValue>("loc.men_postalcode").Value.ToString()}</p>";
                }
            }
            return text;
        }
        public static string TermsandCondition(IOrganizationService service, ITracingService tracingService)
        {
            string text = string.Empty;
            tracingService.Trace("CourseOrderConfirmation: Start of T&C function ");
            string termsandcondition = string.Format(@"<fetch top='1'>
                                                          <entity name='men_mentorsettings'>
                                                            <attribute name='men_termsandconditions' />
                                                          </entity>
                                                        </fetch>");
            EntityCollection tcEntity = service.RetrieveMultiple(new FetchExpression(termsandcondition));

            foreach (Entity e in tcEntity.Entities)
            {
                text = e["men_termsandconditions"].ToString();
            }
            return text;
        }
        public static string GetCourseRequirements(Guid crID, string title, IOrganizationService service, ITracingService tracingService)
        {
            string crInformation = string.Empty;
            string crQuery = string.Format(@"<fetch>
                                              <entity name='men_courserequirement'>
                                                <attribute name='men_liftheight' />
                                                <attribute name='men_spacerequired' />
                                                <filter>
                                                  <condition attribute='men_courserequirementid' operator='eq' value='{0}' />
                                                </filter>
                                                <link-entity name='men_courserequirementdetail' from='men_courserequirementid' to='men_courserequirementid' alias='crDetail'>
                                                  <order attribute='men_sequence' />
                                                  <link-entity name='men_courserequirementtext' from='men_courserequirementtextid' to='men_courserequirementtextid' alias='crText'>
                                                    <attribute name='men_information' />
                                                  </link-entity>
                                                </link-entity>
                                              </entity>
                                            </fetch>", crID.ToString());
            EntityCollection courseReqCollection = service.RetrieveMultiple(new FetchExpression(crQuery));
            crInformation += string.Format(@"<p><strong>Course: {0} </strong></p>", title);
            foreach (Entity e in courseReqCollection.Entities)
            {
                string crInfoText = $"{e.GetAttributeValue<AliasedValue>("crText.men_information").Value.ToString()}<br/><br/>";
                crInformation += crInfoText.Replace("{LiftHeight}", (e["men_liftheight"].ToString())).Replace("{SpaceRequired}", (e["men_spacerequired"].ToString()));
            }
            crInformation += String.Format("</p><br/><p>&nbsp;</p>");
            return crInformation;
        }
        public static Guid GetCourseOrder(Guid quoteID, IOrganizationService service, ITracingService tracingService)
        {
            Guid cOid = Guid.Empty;
            string cOrderQuery = string.Format(@"<fetch>
                                                  <entity name='quote'>
                                                    <attribute name='men_courseorderid' />
                                                    <filter>
                                                      <condition attribute='quoteid' operator='eq' value='{0}' />
                                                    </filter>
                                                  </entity>
                                                </fetch>", quoteID.ToString());
            EntityCollection courseOrder = service.RetrieveMultiple(new FetchExpression(cOrderQuery));
            foreach (Entity e in courseOrder.Entities)
            {
                if (e.Contains("men_courseorderid"))
                {
                    cOid = new Guid(((EntityReference)e["men_courseorderid"]).Id.ToString());
                    tracingService.Trace("CourseorderID", cOid.ToString());
                }

            }
            return cOid;
        }
    }
}
