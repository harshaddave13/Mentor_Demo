using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
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
    "men_regenerateorderconfirmation",
    "Mentor.PlugIns.CourseOrderConfirmation: Update of men_courseorder",
    1000,
    IsolationModeEnum.Sandbox
    , Image1Type = ImageTypeEnum.PreImage
    , Image1Name = "PreImage"
    , Image1Attributes = "ownerid,statuscode"
)]
    public class CourseOrderConfirmation : PluginBase
    {
        protected override void ExecuteCDSPlugin(LocalPluginContext localcontext)
        {
            localcontext.Trace("CourseOrderConfirmation: Start");
            var service = localcontext.OrganizationService;
            var tracingService = localcontext.TracingService;
            tracingService.Trace("CourseOrderConfirmation: Start");
            try
            { 
                Guid courseOrderid = localcontext.PluginExecutionContext.PrimaryEntityId;
                var courseOrderEntity = localcontext.Target.ToEntity<men_courseorder>();
                var preCourseOrder = localcontext.MergedPreTarget.ToEntity<men_courseorder>();               
                
                tracingService.Trace("CourseOrderConfirmation: Converting Target to early bound");
                if (preCourseOrder.statuscode.Value == men_courseorder_statuscode.HardBooking)
                {
                    men_courseorder courseorder = courseOrderEntity;
                    tracingService.Trace("CourseOrderConfirmation: Organisation service");
                    
                    Guid cOid = courseorder.Id;
                    string emailHTML = string.Empty;
                    string coursename = string.Empty;
                    string courseRequirementText = string.Empty;
                    List<string> cOrderInfo = new List<string>();
                    CourseOrderEmailFunctions.GetAccountInfo(cOid, cOrderInfo, service, tracingService);
                    Guid senderID = new Guid(cOrderInfo[5]);
                    Guid toEmailAddress = new Guid(cOrderInfo[4]);
                    string billingInfo = cOrderInfo[0];
                    string siteTrainingLoc = cOrderInfo[2];
                    string greetings = cOrderInfo[1];
                    string csTC = cOrderInfo[3];
                    string starttime = cOrderInfo[7];
                    string standardTC = string.Empty;
                    string quoteid = string.Empty;
                    standardTC = CourseOrderEmailFunctions.TermsandCondition(service, tracingService);
                    tracingService.Trace("CourseOrderConfirmation:Initial HTML with greetings " + senderID.ToString());
                    emailHTML += String.Format(@"<p> <img style='float: right;' 
                                            src='https://mentortraining.co.uk/wp-content/uploads/2022/06/MENTOR-Logo-200px.png' alt='' width='265' height='84' /></p>                                
                                        <p>{0}</p>
                                        <p> Thank you for placing your order with Mentor.</p>
                                        <p><strong>Your Course Order Number {1}</strong></p>
                                        <table border='1'>
                                        <tbody>
                                            <tr>
                                                <td style='width: 284px;'>
                                                    <p> <strong>Billing Information:</strong></p>
                                                </td>
                                                <td style='width: 282px;'>
                                                    <p> <strong>Training Location:</strong></p>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style='width: 284px;'>
                                                    {2}
                                                </td>", greetings, cOrderInfo[6], billingInfo);
                    tracingService.Trace("CourseOrderConfirmation: Greeting HTML " + emailHTML);
                    string top1quote = string.Format(@"<fetch top='1'>
                                                          <entity name='quote'>
                                                            <attribute name='quoteid' />
                                                            <filter type='and'>
                                                              <condition attribute='men_courseorderid' operator='eq' value='{0}' />
                                                              <filter type='or'>
                                                                <condition attribute='statecode' operator='eq' value='1' />
                                                                <condition attribute='statecode' operator='eq' value='2' />
                                                              </filter>
                                                            </filter>
                                                          </entity>
                                                        </fetch>", cOid.ToString());
                    EntityCollection quote = service.RetrieveMultiple(new FetchExpression(top1quote));
                    foreach (Entity entity in quote.Entities)
                    {
                        quoteid = entity.Id.ToString();
                    }
                    string quoteQuery = string.Format(@"<fetch>
                                                          <entity name='quote'>
                                                            <attribute name='totalamount' />
                                                            <filter>
                                                              <condition attribute='quoteid' operator='eq' value='{0}' />
                                                            </filter>
                                                            <link-entity name='quotedetail' from='quoteid' to='quoteid' link-type='outer' alias='QD'>
                                                              <attribute name='baseamount' />
                                                              <attribute name='men_mentorproductype' />
                                                              <order attribute='sequencenumber' />
                                                              <link-entity name='product' from='productid' to='productid' link-type='outer' alias='Prod'>
                                                                <attribute name='name' />
                                                              </link-entity>
                                                              <link-entity name='men_course' from='men_courseid' to='men_courseid' link-type='outer' alias='Course'>
                                                                <attribute name='men_coursedatesoverride' />
                                                                <attribute name='men_courseid' />
                                                                <attribute name='men_coursetitle' />
                                                                <attribute name='men_numberofdelegates' />
                                                                <attribute name='men_trainingcentrecourseid' />
                                                                <link-entity name='men_experiencelevel' from='men_experiencelevelid' to='men_experiencelevelid' link-type='outer' alias='ExLvl'>
                                                                  <attribute name='men_name' />
                                                                </link-entity>
                                                                <link-entity name='men_pricebook' from='men_pricebookid' to='men_pricebookid' link-type='outer' alias='PB'>
                                                                  <attribute name='men_courserequirementid' />
                                                                </link-entity>
                                                              </link-entity>
                                                            </link-entity>
                                                          </entity>
                                                        </fetch>", quoteid);
                    EntityCollection entities = service.RetrieveMultiple(new FetchExpression(quoteQuery));

                    for (int i = 0; i < entities.Entities.Count; i++)
                    {
                        Entity e = entities.Entities[i];
                        Guid courseId = Guid.Empty;
                        OptionSetValue prodType = null;
                        string cost = string.Empty;
                        string totalcost = string.Empty;
                        tracingService.Trace("CreateQuoteEmail: Inside loop ");
                        if (e.Contains("Course.men_courseid"))
                        {
                            courseId = new Guid(e.GetAttributeValue<AliasedValue>("Course.men_courseid").Value.ToString());
                            tracingService.Trace("CreateQuoteEmail: Courseid retrieved " + courseId);
                        }
                        if (e.Contains("QD.men_mentorproductype"))
                        {
                            prodType = (OptionSetValue)e.GetAttributeValue<AliasedValue>("QD.men_mentorproductype").Value;
                            tracingService.Trace("CreateQuoteEmail: producttype " + prodType.Value.ToString());
                        }
                        if (e.Contains("QD.baseamount"))
                        {
                            cost = ((Money)e.GetAttributeValue<AliasedValue>("QD.baseamount").Value).Value.ToString("0.00");
                        }

                        if (e.Contains("totalamount"))
                        {
                            totalcost = ((Money)e["totalamount"]).Value.ToString("0.00");
                        }
                        if (e.Contains("Course.men_coursetitle"))
                        {
                            coursename = e.GetAttributeValue<AliasedValue>("Course.men_coursetitle").Value.ToString();
                        }
                        if (prodType != null && prodType.Value == 1)
                        {
                            string expLvl = string.Empty;
                            string nofDelegates = string.Empty;
                            string trainingdates = string.Empty;

                            tracingService.Trace("CourseOrderConfirmation: product type is training ");
                            if (i == 0)
                            {
                                tracingService.Trace("CourseOrderConfirmation: 1st loop ");
                                if (e.Contains("Course.men_trainingcentrecourseid"))
                                {
                                    tracingService.Trace("CourseOrderConfirmation: course is tcc ");
                                    string traininglocation = CourseOrderEmailFunctions.GetTrainingLocation(courseId, service, tracingService);

                                    emailHTML += String.Format(@"<td style='width: 284px;'>
                                                            {0}                                                           
                                                        </td>
                                                        </tr>
                                                        </tbody>
                                                    </table>
                                                    <p> &nbsp;</p>
                                                    <p><strong>Course Details:</strong></p>
                                                        <table border='1'>
                                                            <tbody>
                                                                <tr>
                                                                    <td>
                                                                        <p><strong>Course</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Experience Level</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>No. of Delegates</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Training Date(s)</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Start Times</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Cost</strong></p>
                                                                    </td>
                                                                </tr>", traininglocation);
                                }
                                else
                                {
                                    emailHTML += String.Format(@"<td style='width: 284px;'>
                                                            {0}  
                                                        </td>
                                                        </tr>
                                                        </tbody>
                                                        </table>
                                                    <p> &nbsp;</p>
                                                        <p><strong>Course Details:</strong></p>
                                                        <table border='1'>
                                                            <tbody>
                                                                <tr>
                                                                    <td>
                                                                        <p><strong>Course</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Experience Level</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>No. of Delegates</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Training Date(s)</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Start Times</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Cost</strong></p>
                                                                    </td>
                                                                </tr>", siteTrainingLoc);
                                }
                                tracingService.Trace("CourseOrderConfirmation: add training location to HTML on the first loop");
                            }
                            else if ((i < entities.Entities.Count - 1) && i > 0)
                            {
                                emailHTML += String.Format(@"</tbody>
                                                    </table>
                                                        <p><strong>Course Details:</strong></p>
                                                        <table border='1'>
                                                            <tbody>
                                                                <tr>
                                                                    <td>
                                                                        <p><strong>Course</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Experience Level</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>No. of Delegates</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Training Date(s)</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Start Times</strong></p>
                                                                    </td>
                                                                    <td>
                                                                        <p><strong>Cost</strong></p>
                                                                    </td>
                                                                </tr>");

                                tracingService.Trace("CourseOrderConfirmation: Set Variable course title ");
                            }
                            if (e.Contains("ExLvl.men_name"))
                            {
                                expLvl = $"<p> {e.GetAttributeValue<AliasedValue>("ExLvl.men_name").Value.ToString()} </p>";
                            }
                            if (e.Contains("Course.men_numberofdelegates"))
                            {
                                nofDelegates = $"<p> {e.GetAttributeValue<AliasedValue>("Course.men_numberofdelegates").Value.ToString()}</p>";
                            }

                            if (e.Contains("Course.men_coursedatesoverride"))
                            {
                                trainingdates = $"<p> {e.GetAttributeValue<AliasedValue>("Course.men_coursedatesoverride").Value.ToString()}</p>";
                            }

                            emailHTML += String.Format(@"<tr>
                                                <td width='140'>
                                                <p>{0}</p>
                                                </td>
                                                <td width='103' align='center' >
                                                    {1}
                                                </td>
                                                <td width='79' align='center'>
                                                <p>{2}</p>
                                                </td>
                                                <td width='100' align='center'>
                                                <p>{3}</p>
                                                </td>
                                                <td width='52' align='center'>
                                                <p>{4}</p>
                                                </td>
                                                <td width='59' align='center'>
                                                <p>&pound;{5}</p>
                                                </td>
                                                </tr>", coursename, expLvl, nofDelegates, trainingdates, starttime, cost);


                            if (e.Contains("PB.men_courserequirementid"))
                            {
                                tracingService.Trace("CourseOrderConfirmation:set courserequirement Guid" + e.GetAttributeValue<AliasedValue>("PB.men_courserequirementid").Value);
                                var courseRequirement = (EntityReference)e.GetAttributeValue<AliasedValue>("PB.men_courserequirementid").Value;
                                Guid courseRequirementID = courseRequirement.Id;
                                courseRequirementText += CourseOrderEmailFunctions.GetCourseRequirements(courseRequirementID, coursename, service, tracingService);
                            }
                        }
                        else
                        {
                            if (e.Contains("Prod.name"))
                            {
                                emailHTML += String.Format(@"<tr>
                                                <td width='140'>
                                                <p>{0}</p>
                                                </td>
                                                <td width='103' align='center' >
                                                    <p>&nbsp;</p>
                                                </td>
                                                <td width='79' align='center'>
                                                <p>&nbsp;</p>
                                                </td>
                                                <td width='100' align='center'>
                                                <p>&nbsp;</p>
                                                </td>
                                                <td width='52' align='center'>
                                                <p>&nbsp;</p>
                                                </td>
                                                <td width='59' align='center'>
                                                <p>&pound;{1}</p>
                                                </td>
                                                </tr>", e.GetAttributeValue<AliasedValue>("Prod.name").Value.ToString(), cost);
                            }

                        }
                        if (i == entities.Entities.Count - 1)
                        {
                            emailHTML += String.Format(@"</tbody>
                                                    </table>
                                            <p><strong>Total Cost (Ex VAT):&nbsp;&nbsp;&nbsp; &pound;{0}</strong></p>

                                            <p>&nbsp;</p>
                                                
                                            <p><strong>In accepting this booking, you have agreed with our terms &amp; conditions.</strong></p>", totalcost);
                        }

                    }
                    if (courseRequirementText != string.Empty)
                    {
                        emailHTML += String.Format(@"<p>&nbsp;</p><p><strong>What we need:</strong></p>{0}", courseRequirementText);
                    }
                    emailHTML += string.Format("<p><strong>Terms and Conditions </strong></p>");
                    if (csTC != string.Empty)
                    {
                        emailHTML += string.Format(csTC);
                    }
                    else
                    {
                        emailHTML += string.Format(standardTC);
                    }
                    emailHTML += string.Format("<p>&nbsp;</p><p><strong> Thank you for your business</strong> </p> ");
                    tracingService.Trace(emailHTML);

                    List<ActivityParty> fromEmailList = new List<ActivityParty>();
                    ActivityParty fromActivityParty = new ActivityParty();
                    fromActivityParty.PartyId = new EntityReference("systemuser", senderID);
                    fromEmailList.Add(fromActivityParty);

                    List<ActivityParty> toEmailList = new List<ActivityParty>();
                    ActivityParty toActivityParty = new ActivityParty();
                    toActivityParty.PartyId = new EntityReference("contact", toEmailAddress);
                    toEmailList.Add(toActivityParty);

                    Email email = new Email()
                    {
                        Description = emailHTML,
                        From = fromEmailList,
                        DeliveryReceiptRequested = false,
                        OwnerId = new EntityReference("systemuser", senderID),
                        IsBilled = false,
                        ReadReceiptRequested = false,
                        RegardingObjectId = new EntityReference(men_courseorder.EntityLogicalName, courseorder.Id),
                        DirectionCode = true,
                        Subject = "Order Confirmation for Training",
                        To = toEmailList
                    };
                    service.Create(email);
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace("CourseOrderConfirmation: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the CourseOrderConfirmation plugin " + ex.Message, ex);
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace("CourseOrderConfirmation: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("A standard exception error occurred in the CourseOrderConfirmation plugin " + ex.Message, ex);
            }

            tracingService.Trace("CourseOrderConfirmation: End");
        }
    }
}
