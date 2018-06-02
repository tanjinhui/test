using iTextSharp.text;
using iTextSharp.text.pdf;
using Jayrock.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using TRIMSDK;
using HPSDK = HP.HPTRIM.SDK;
using System.Diagnostics;
using System.Reflection;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SaffronWebApp.Web
{
    public partial class ConsignmentsModule : BasepageSessionExpire
    {
        public Database db;
        private string zipFolder = "C:\\erks\\Export\\Consignment";
        private const string CONSIGNMENT_FOLDER = "Consignment";
        private static readonly string indexUrl = System.Configuration.ConfigurationManager.AppSettings["WebSite"]; 
        private static readonly string _stubReportPath = System.Configuration.ConfigurationManager.AppSettings["StubReportPath"];
        private static readonly string _sendMailToApprover = System.Configuration.ConfigurationManager.AppSettings["SendMailToApprover"];
        private static readonly string _sendMailToNewConsignment = System.Configuration.ConfigurationManager.AppSettings["SendMailToNewConsignment"];

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string action = Request["action"] + "";
                GlobalFunc.LogDebug(action + " action is called ");

                switch (action.ToLower())
                {
                    case "getallconsignments":
                        GetAllConsignmentsHPE();
                        //GetAllConsignments();
                        break;
                    case "getissuesbyconsignment":
                        GetConsignmentIssuesByConsignmentNumber();
                        break;
                    case "getapprovalbyconsignment":
                        GetApprovalByConsignment();
                        break;
                    case "save":
                        SaveConsignmentHPE();
                        //SaveConsignment();
                        break;
                    case "getconsignment":
                        GetConsignmentHPE();
                        break;
                    case "removeformconsignment":
                        RemoveFromConsignment();
                        break;
                    case "returntohome":
                        ReturnToHome();
                        break;
                    case "encloseitem":
                        EncloseItem();
                        break;
                    case "showdetails":
                        ShowDetail();
                        break;
                    case "getsecuritaudit":
                        GetSecuritAudit();
                        break;
                    case "completereview":
                        CompleteReview();
                        break;
                    case "getparts":
                        GetParts();
                        break;
                    case "sendemail":
                        SendMailToApprover();
                        break;
                    case "checkonhold":
                        CheckOnHoldRecord();
                        break;
                    case "complete":
                        CompleteConsignment();
                        break;
                    case "logreport":
                        var Method = Request["method"];
                        var consignmentNo = Request["consignmentNo"];
                        HPSDK.Record Review = null;
                        HPSDK.Consignment consignment = null;
                        using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                        {
                            CheckDiffConsignmentWithChild(database, consignmentNo);
                            if (Method != "Review")
                                consignment = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));
                            else
                                Review = new HPSDK.Record(database, Convert.ToInt32(consignmentNo));
                            Report(database, consignmentNo, Method, consignment, Review, "");
                        }

                        //ModifyRecordType(consignmentNo, "");
                        break;
                    case "consignmentexport":
                        ConsignmentExport(Request["consignmentNo"]);
                        break;
                    case "getloginname":
                        GetLoginName();
                        break;
                    case "exportissues":
                        ExportIssues();
                        break;
                    case "download":
                        DownloadFile();
                        break;
                    case "delete":
                        DeleteConsignment();
                        break;
                    case "consignmentdetailsexport":
                        ConsignmentDetailsExport();
                        break;
                    case "checkaccess":
                        CheckConsignmentAccessByUser();
                        break;
                    case "reexamine":
                        ReexamineIssues();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            // finally
            // {
            Response.End();
            //  }
        }


        private void GetParts()
        {
            try
            {
                string uri = Request.Form.Get("Uri") + "";
                string consignmentNo = Request.Form.Get("consignmentNo") + "";
                string code = Request.Form.Get("Code") + "";
                string treetype = Request.Form.Get("type") + "";
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartArray();
                    int count = 0;
                    bool showText = false;
                    using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                    {
                        string searchstring = "";
                        var conslist = new HPSDK.TrimURIList();
                        if (uri != "source" && uri != "")
                        {
                            //Parent = new HPSDK.Schedule(database, Convert.ToInt32(uri));
                            searchstring = $"container:\"{code}\"";
                            conslist = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                        }
                        else
                        {
                            showText = true;
                            string strNot = treetype == "NonElectronic" ? "not FolderType:Electronic and" : "";
                            //searchstring = $"consignment:{consignmentNo} AND {strNot} type:Part";
                            //conslist = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                            //if (conslist.Count == 0)
                            //{
                            searchstring = $"consignment:{consignmentNo} AND {strNot} type:Part";
                            conslist = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                            //foreach (var cons in consList)
                            //{
                            //    var recsearchstring = $"Uri:{cons.ToString()}+ AND {strNot} type:Part "; // and not closedOn:blank";
                            //    var reclist = Search(database, HPSDK.BaseObjectTypes.Record, recsearchstring);
                            //    foreach (var recuri in reclist)
                            //    {
                            //        if (!conslist.Contains(recuri))
                            //            conslist.Add(recuri);
                            //    }
                            //}
                            //}
                        }
                        bool Leaf = true;
                        foreach (var itemUri in conslist)
                        {
                            var rec = new HPSDK.Record(database, Convert.ToInt32(itemUri));
                            JW.WriteStartObject();
                            if (rec.Contents != "")
                            {
                                Leaf = false;
                            }
                            else
                            {
                                Leaf = true;
                            }

                            JW.WriteMember("Uri");
                            JW.WriteString(rec.Uri.ToString());

                            JW.WriteMember("Classificationcode");
                            string classcode = GetFieldValue(database, "Classification code", rec);
                            JW.WriteString(classcode == "" ? rec.Number : classcode);

                            string classpath = GetFieldValue(database, "Classification path", rec);
                            JW.WriteMember("Classificationpath");
                            JW.WriteString(classpath == "" ? rec.FilePath : classpath);
                            JW.WriteMember("Title");
                            JW.WriteString(rec.Title);

                            if (treetype == "Electronic")
                            {
                                JW.WriteMember("DAnumber");
                                string DAno = GetFieldValue(database, "GRS deposit identifier", rec);
                                JW.WriteString(DAno);
                                JW.WriteMember("showText");
                                JW.WriteBoolean(showText);
                            }

                            JW.WriteMember("number");
                            JW.WriteString(rec.Number);

                            string Icon = "";
                            if (rec.IsElectronic)
                            {
                                //Icon = "attachment";
                                Icon = GetIcon(ref rec, rec.Extension.ToLower());
                            }
                            else
                            {
                                Icon = GetRecordIcon(rec.RecordType);
                            }
                            JW.WriteMember("iconCls");
                            JW.WriteString(Icon);

                            String FolderTypeIcon = GetFolderIcon(rec, database);

                            JW.WriteMember("FolderTypeIcon");
                            JW.WriteString(FolderTypeIcon);

                            JW.WriteMember("SecurityClassification");
                            JW.WriteString(rec.Security.ToString());

                            JW.WriteMember("leaf");
                            JW.WriteBoolean(Leaf);

                            JW.WriteEndObject();
                        }
                    }
                    JW.WriteEndArray();
                }

            }
            catch (Exception ex)
            {

            }
        }

        private bool WriteIssueJson(JsonWriter JW, HPSDK.ConsignmentIssue issue)
        {
            try
            {
                //var recordTitle = issue.Name == "Access denied." ? "" : issue.Record.Name ;
                JW.WriteStartObject();
                JW.WriteMember("IssueType");

                JW.WriteString(issue.IssueType.ToString());
                JW.WriteMember("IssueRecord");
                JW.WriteString(issue.Name);
                JW.WriteMember("Description");
                JW.WriteString(issue.Description);
                JW.WriteEndObject();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private void WritePartJson(HPSDK.Database database, JsonWriter writer, long uri)
        {

            var aRecord = new HPSDK.Record(database, new HPSDK.TrimURI(uri));
            var isPartLevel = aRecord.TrimType == HPSDK.BaseObjectTypes.Consignment;
            var consignmentCondition = "container[consignment: " + uri.ToString() + "]";
            var searchString = (isPartLevel ? "PartNumber:\"\" and Type:Part" + " AND " + consignmentCondition : "container[uri:" + uri.ToString() + "]");//part.Uri.Value.ToString();

            var recordUris = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchString).ToList();
            //if (!isPartLevel && recordUris.Any()) {
            //    writer.WriteMember("children");
            //}
            //if (recordUris.Any())
            //{
            writer.WriteStartArray();
            foreach (var recordUri in recordUris)
            {
                var record = new HPSDK.Record(database, recordUri);
                writer.WriteStartObject();
                writer.WriteMember("Uri");
                writer.WriteNumber(record.Uri.Value.ToString());
                writer.WriteMember("leaf");
                writer.WriteBoolean(!HasChild(database, record.Uri));
                writer.WriteMember("description");
                writer.WriteString(record.Title);
                writer.WriteMember("recordNumber");
                writer.WriteString(record.Number);
                if (isPartLevel)
                {
                    writer.WriteMember("Classificationcode");
                    writer.WriteString(record.GetFieldValue(new HPSDK.FieldDefinition(database, "Classification code")).ToString());
                    writer.WriteMember("Classificationpath");
                    writer.WriteString(record.GetFieldValue(new HPSDK.FieldDefinition(database, "Classification path")).ToString());
                }
                //writer.WriteMember("path");
                //writer.WriteString(record.Classification.co)


                //WritePartJson(database, writer, record.Uri);
                writer.WriteEndObject();
                //WritePartJson(database, writer, part);
            }
            writer.WriteEndArray();
            //}
        }
        private bool HasChild(HPSDK.Database database, long uri)
        {
            var searchString = "container[uri:" + uri.ToString() + "]";//part.Uri.Value.ToString();

            var recordUris = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchString).ToList();

            return recordUris.Any();
        }

        #region sendmail
        private void SendMailToApprover()
        {
            try
            {
                var consignmentNo = Request["consignmentNo"];
                var consignmentName = Request["consignmentName"];
                var method = Request["method"];

                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                {
                    var approvers = GetApproverEmailAddressByConsignmentUri(database, consignmentNo, method);
                    if (approvers != null && approvers.Count > 0)
                    {
                        var emailTemplate = new HPSDK.Record(database, _sendMailToApprover);
                        var linkUrl = string.Format("{0}?HPERMURI={1}CNO{2}", indexUrl, consignmentNo, method);
                        foreach (var address in approvers)
                        {
                            SendMail(database, consignmentName, address, emailTemplate, linkUrl);

                            if (method.ToLower() == "review")
                            {
                                var review = new HPSDK.Record(database, Convert.ToInt32(consignmentNo));
                                review.SetFieldValue(new HPSDK.FieldDefinition(database, "Email Sent"), new HPSDK.UserFieldValue(true));
                                review.SetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Sent"), new HPSDK.UserFieldValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                                review.Save();
                            }
                            else
                            {
                                var consi = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));
                                consi.SetFieldValue(new HPSDK.FieldDefinition(database, "Email Sent"), new HPSDK.UserFieldValue(true));
                                consi.SetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Sent"), new HPSDK.UserFieldValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                                consi.Save();
                            }
                        }
                        using (JsonWriter JW = CreateJsonWriter(Response.Output))
                        {
                            JW.WriteStartObject();
                            JW.WriteMember("success");
                            JW.WriteBoolean(true);
                            JW.WriteMember("message");
                            JW.WriteString("");
                            JW.WriteEndObject();
                        }
                    }
                    else {
                        using (JsonWriter JW = CreateJsonWriter(Response.Output))
                        {
                            JW.WriteStartObject();
                            JW.WriteMember("success");
                            JW.WriteBoolean(false);
                            JW.WriteMember("message");
                            JW.WriteString("Approver is null");
                            JW.WriteEndObject();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString("");
                    JW.WriteEndObject();
                }
                GlobalFunc.Log(ex);
            }
        }

        private void SendMail(HPSDK.Database database, string consignmentName, string address, HPSDK.Record emailTemplate, string linkUrl)
        {
            if (consignmentName != "" && address != "" && emailTemplate != null)
            {
                var emailRecord = new HPSDK.Record(database, new HPSDK.RecordType(database, "Send Email"));
                HPSDK.FieldDefinition subjectDef = new HPSDK.FieldDefinition(database, "Email Subject");
                HPSDK.FieldDefinition receiverDef = new HPSDK.FieldDefinition(database, "Recipient email");
                HPSDK.FieldDefinition bodyDef = new HPSDK.FieldDefinition(database, "Email Content");
                HPSDK.FieldDefinition isHtmlDef = new HPSDK.FieldDefinition(database, "Is Html Body");
                //var emailTemplate = new HPSDK.Record(database, "EMT/18/1");
                var contentTemplate = emailTemplate.GetFieldValue(bodyDef).AsString();
                var isHtml = emailTemplate.GetFieldValue(isHtmlDef).AsBool();
                var subject = emailTemplate.GetFieldValue(subjectDef).AsString();
                emailRecord.Title = string.Format("{0}:{1}", subject, consignmentName);
                emailRecord.SetFieldValue(subjectDef, new HPSDK.UserFieldValue(subject));
                emailRecord.SetFieldValue(receiverDef, new HPSDK.UserFieldValue(address));
                //emailRecord.SetFieldValue(receiverDef, new HPSDK.UserFieldValue("by.su@gti.com.hk"));
                var content = contentTemplate.Replace("%_LINK_%", linkUrl).Replace("%_CONSIGNMENT_NUMBER_%", consignmentName);

                emailRecord.SetFieldValue(bodyDef, new HPSDK.UserFieldValue(content));
                emailRecord.SetFieldValue(isHtmlDef, new HPSDK.UserFieldValue(isHtml));
                emailRecord.Save();
            }
        }
        
        private List<string> GetApproverEmailAddressByConsignmentUri(HPSDK.Database database, string consignmentNo, string method)
        {
            var approvalList = new List<string>();
            HPSDK.TrimURIList uriList = null;
            var searchstring = "";
            if (method.ToLower() == "review")
            {
                searchstring = $"ConsignmentReviewUri:{consignmentNo}";
                uriList = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                if (uriList.Count > 0)
                {
                    foreach (var approvaluri in uriList)
                    {
                        var approval = new HPSDK.Record(database, approvaluri);
                        if (approval != null && approval.Name != "")
                        {
                            approvalList.Add(approval.OwnerLocation.EmailAddress);
                        }
                    }
                }
            }
            else
            {
                searchstring = $"consignment:{consignmentNo}";
                uriList = Search(database, HPSDK.BaseObjectTypes.ConsignmentApprover, searchstring);
                if (uriList.Count > 0)
                {
                    foreach (var approvaluri in uriList)
                    {
                        var approval = new HPSDK.ConsignmentApprover(database, approvaluri);
                        if (approval != null && approval.Name != "")
                        {
                            approvalList.Add(approval.Approver.EmailAddress);
                        }
                    }
                }
            }
            return approvalList;
        }
        #endregion
        
        private void GetAllConsignmentsHPE()
        {
            try
            {
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
                {

                    var consList = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Consignment, "all");
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();

                        JW.WriteMember("results");
                        JW.WriteStartArray();
                        foreach (var consUri in consList)
                        {
                            HPSDK.Consignment consignment = new HPSDK.Consignment(database, Convert.ToInt32(consUri));
                            //var record = 0;
                            //var issues = 0;

                            //    var listRecord = SearchShowDetails(database, consignment.Number);
                            //    record = listRecord == null ? 0 : listRecord.Count;

                            //    var listIssues = SearchConsignmentIssue(database, consignment.Number);
                            //    issues = listIssues == null ? 0 : listIssues.Count;

                            JW.WriteStartObject();
                            JW.WriteMember("Uri");
                            JW.WriteString(consignment.Uri.ToString());
                            JW.WriteMember("ConsignmentNumber");
                            JW.WriteString(consignment.Number);

                            JW.WriteMember("Description");
                            JW.WriteString(consignment.Description);

                            string disposalmethod = consignment.GetFieldValue(new HPSDK.FieldDefinition(database, "ConsignmentMethod")).ToString();
                            JW.WriteMember("DisposalMethod");
                            JW.WriteString(string.IsNullOrEmpty(disposalmethod) ? consignment.DisposalMethod.ToString() : disposalmethod);
                            JW.WriteMember("CutoffDate");
                            JW.WriteString(consignment.CutoffDate.ToShortDateString());

                            int recordCount = Search(database, HPSDK.BaseObjectTypes.Record, $"ConsignmentUriFlag:{consUri}").Count;
                            JW.WriteMember("ItemsInConsignment");
                            JW.WriteString(recordCount.ToString());

                            int issuesCount = SearchConsignmentIssue(database, consUri.ToString()).Count;
                            JW.WriteMember("ItemsWithIssues");
                            JW.WriteString(issuesCount.ToString());
                            JW.WriteMember("CurrentStatus");
                            JW.WriteString(consignment.StateDescription);
                            try
                            {
                                string strConversion = consignment.GetFieldValue(new HPSDK.FieldDefinition(database, "Is Conversion")).ToString();
                                JW.WriteMember("IsConversion");
                                JW.WriteString(strConversion);
                            }
                            catch (Exception)
                            {
                                JW.WriteMember("IsConversion");
                                JW.WriteString("No");
                            }

                            JW.WriteMember("IsReview");
                            JW.WriteBoolean(false);
                            JW.WriteMember("EmailSent");
                            JW.WriteBoolean(consignment.GetFieldValue(new HPSDK.FieldDefinition(database, "Email Sent")).AsBool());
                            JW.WriteMember("IsComplete");
                            JW.WriteBoolean(consignment.IsComplete);

                            JW.WriteEndObject();
                        }

                        //Get Review Consignment

                        var reviewList = Search(database, HPSDK.BaseObjectTypes.Record, "type:[default:\"Review consignment\"]");
                        foreach (var reviewUri in reviewList)
                        {
                            var reviewRec = new HPSDK.Record(database, Convert.ToInt32(reviewUri));
                            JW.WriteStartObject();
                            JW.WriteMember("Uri");
                            JW.WriteString(reviewRec.Uri.ToString());
                            JW.WriteMember("ConsignmentNumber");
                            JW.WriteString(reviewRec.Title);

                            string description = GetFieldValue(database, "Description", reviewRec);
                            JW.WriteMember("Description");
                            JW.WriteString(description);

                            string disposalmethod = reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "ConsignmentMethod")).ToString();
                            JW.WriteMember("DisposalMethod");
                            JW.WriteString(string.IsNullOrEmpty(disposalmethod) ? reviewRec.DisposalMethod.ToString() : disposalmethod);

                            try
                            {
                                string cutoffdate = reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment CutOffDate")).AsDate().ToShortDateString();
                                JW.WriteMember("CutoffDate");
                                JW.WriteString(cutoffdate);
                            }
                            catch
                            {
                                JW.WriteMember("CutoffDate");
                                JW.WriteString("");
                            }

                            var iteminconsignment = Search(database, HPSDK.BaseObjectTypes.Record, $"ConsignmentReviewUri:{reviewUri.ToString()}");//GetFieldValue(database, "ConsignmentReviewRecordsCount", reviewRec);
                            JW.WriteMember("ItemsInConsignment");
                            JW.WriteString(iteminconsignment.Count.ToString());
                            JW.WriteMember("ItemsWithIssues");
                            JW.WriteString("0");

                            string CurrentStatus = GetFieldValue(database, "Consignment Status", reviewRec);
                            JW.WriteMember("CurrentStatus");
                            JW.WriteString(CurrentStatus);
                            try
                            {
                                string strConversion = reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "Is Conversion")).ToString();
                                JW.WriteMember("IsConversion");
                                JW.WriteString(strConversion);
                            }
                            catch (Exception ex)
                            {
                                string strConversion = reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "Is Conversion")).ToString();
                                JW.WriteMember("IsConversion");
                                JW.WriteString(strConversion);
                            }

                            try
                            {
                                var datereview = reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Review Date")).AsDate();
                            }
                            catch (Exception)
                            {

                            }

                            JW.WriteMember("IsReview");
                            JW.WriteBoolean(true);
                            JW.WriteMember("EmailSent");
                            JW.WriteBoolean(reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "Email Sent")).AsBool());
                            JW.WriteMember("IsComplete");
                            JW.WriteBoolean(true);
                            //JW.WriteBoolean(reviewRec.DateApproved.Year != -1);

                            JW.WriteEndObject();
                        }

                        JW.WriteEndArray();

                        JW.WriteMember("totalrows");
                        JW.WriteNumber(consList.Count);

                        JW.WriteEndObject();
                    }
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString("");
                    JW.WriteEndObject();
                }
                GlobalFunc.Log(ex);
            }
        }

        private HP.HPTRIM.SDK.TrimURIList Search(HP.HPTRIM.SDK.Database database, HP.HPTRIM.SDK.BaseObjectTypes baseObjectTypes, string searchString, string udfSortField = "")
        {
            var trimMainObjectSearch = new HP.HPTRIM.SDK.TrimMainObjectSearch(database, baseObjectTypes);
            trimMainObjectSearch.SetSearchString(searchString);
            HPSDK.TrimURIList uriArray = null;
            if (!string.IsNullOrWhiteSpace(udfSortField))
            {
                var fieldDefinition = new HPSDK.FieldDefinition(database, udfSortField);
                if (fieldDefinition != null)
                {
                    var propertyOrFieldDef = new HPSDK.PropertyOrFieldDef(fieldDefinition);
                    uriArray = trimMainObjectSearch.GetResultAsUriArraySorted(propertyOrFieldDef, true);
                }
            }
            else
                uriArray = trimMainObjectSearch.GetResultAsUriArray();
            return uriArray;
        }

        private HP.HPTRIM.SDK.TrimURIList SearchConsignmentIssue(HP.HPTRIM.SDK.Database database, string consignmentNumber, string filter = "")
        {
            string searchstring = $"consignment:\"{consignmentNumber}\"";
            if (filter != "")
                searchstring += " and " + filter;
            return Search(database, HP.HPTRIM.SDK.BaseObjectTypes.ConsignmentIssue, searchstring);
        }

        private HPSDK.TrimURIList SearchConsignmentApprovals(HPSDK.Database database, string consignmentNo, string isPending)
        {
            string searchstring = $"consignment:{consignmentNo}";
            if (isPending != "")
                searchstring += " AND status:\"Pending\"";
            return Search(database, HPSDK.BaseObjectTypes.ConsignmentApprover, searchstring);
        }

        private HPSDK.TrimURIList SearchIssuesAggregation(HPSDK.Database database, string containerNo)
        {
            string searchstring = $"container:[default:{containerNo}]";
            //string field = "Container";
            return Search(database, HPSDK.BaseObjectTypes.Record, searchstring);

        }

        private HPSDK.TrimURIList SearchShowDetails(HPSDK.Database database, string consignmentNo)
        {
            string searchstring = $"consignment:{consignmentNo}";
            return Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
        }

        private HPSDK.TrimURIList SearchConsignmentAudit(HPSDK.Database database, string consignmentNo, string method)
        {
            string objecttype = "consignment";
            if (method.ToLower() == "review")
                objecttype = "record";
            string searchstring = $"object:{objecttype},{consignmentNo}";
            return Search(database, HPSDK.BaseObjectTypes.History, searchstring);
        }

        private void GetConsignmentIssuesByConsignmentNumber()
        {
            try
            {
                string consignmentNumber = Request["consignmentNo"];
                bool isReview = Convert.ToBoolean(Request["isReview"]);
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();

                    JW.WriteMember("results");
                    JW.WriteStartArray();
                    var workgroup = System.Configuration.ConfigurationManager.AppSettings["WorkgroupName"];
                    int count = 0;
                    int index = 0;
                    using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
                    {
                        if (!isReview)
                        {
                            var issues = SearchConsignmentIssue(database, consignmentNumber);
                            //db = GlobalFunc.ConnectDBByAdmin;
                            count = issues.Count;
                            foreach (var issueItem in issues)
                            {
                                var issue = new HP.HPTRIM.SDK.ConsignmentIssue(database, issueItem);
                                //auto enclosed
                                //if (issue.Description.Contains("enclosed"))
                                //{
                                //    try
                                //    {
                                //        var rec = new HPSDK.Record(database, issue.Record.Uri);
                                //        rec.IsEnclosed = true;
                                //        rec.Save();
                                //        issue.RemoveIfResolved();
                                //        continue;
                                //    }
                                //    catch
                                //    { }
                                //}

                                JW.WriteStartObject();
                                JW.WriteMember("Title");
                                JW.WriteString(issue.Record.Title);

                                JW.WriteMember("Classificationcode");
                                string classcode = GetFieldValue(database, "Classification code", issue.Record);
                                JW.WriteString(classcode);

                                string classpath = GetFieldValue(database, "Classification path", issue.Record);
                                JW.WriteMember("Classificationpath");
                                JW.WriteString(classpath);
                                string HPEuri = issue.Consignment.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                                if (HPEuri == "")
                                    HPEuri = GlobalFunc.DATASETID + "I" + consignmentNumber;
                                JW.WriteMember("SystemIdentifie");
                                JW.WriteString(HPEuri);
                                JW.WriteMember("Uri");
                                JW.WriteString(issue.Uri.ToString());
                                JW.WriteMember("Description");
                                JW.WriteString(issue.Description);
                                JW.WriteMember("IsElectronic");
                                JW.WriteBoolean(issue.Record.IsElectronic);
                                JW.WriteEndObject();
                                index++;
                                //if (index > 10)
                                //    break;
                            }
                        }
                        JW.WriteEndArray();

                        JW.WriteMember("totalrows");
                        JW.WriteNumber(count);

                        JW.WriteEndObject();
                        //database.Connect(); //Connect with either .ConnectAs or .Connect method

                        //Do something, replace example code here
                    }



                }


            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString("");
                    JW.WriteEndObject();
                }
                GlobalFunc.Log(ex);
            }

        }

        private void ReexamineIssues()
        {
            string consignmentNumber = Request["consignmentNo"];
            string consignmentName = Request["consignmentName"];
            bool isReview = Convert.ToBoolean(Request["isReview"]);
            string issuemessage = "";
            bool hasIssues = false;
            try
            {
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                {
                    var issues = SearchConsignmentIssue(database, consignmentNumber);
                    foreach (var issueItem in issues)
                    {
                        try
                        {
                            var issue = new HP.HPTRIM.SDK.ConsignmentIssue(database, issueItem);
                            if(issue.IssueType != HPSDK.ConsignmentItemIssueType.IssueWarning)
                                issue.RemoveIfResolved();
                        }
                        catch
                        { }
                    }

                    issues = SearchConsignmentIssue(database, consignmentNumber);
                    if (issues.Count > 0)
                    {
                        issuemessage = $"Re-examination of consignment '{consignmentName}' failed. \r\n{issues.Count.ToString()} issues were detected that require resolution.\rDo you want to review these issues now?";
                        hasIssues = true;
                    }
                    else
                    {
                        issuemessage = $"Re-examination of consignment '{consignmentName}' successful.\r\n{issues.Count.ToString()} records examined.\rContinued processing of the consignment can proceed.";
                    }
                }
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(true);
                    JW.WriteMember("hasIssues");
                    JW.WriteBoolean(hasIssues);
                    JW.WriteMember("message");
                    JW.WriteString(issuemessage);
                    JW.WriteEndObject();
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString(ex.Message);
                    JW.WriteEndObject();
                }
            }

        }

        private void GetApprovalByConsignment()
        {
            string consignmentNo = Request["consignmentNo"];
            string Pending = Request["approvaltype"];
            string Method = Request["method"];
            try
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();

                    JW.WriteMember("results");
                    JW.WriteStartArray();
                    int count = 0;
                    int index = 0;
                    using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                    {
                        if (Method.ToLower() != "review")
                        {
                            var approvalList = SearchConsignmentApprovals(database, consignmentNo, Pending);
                            count = approvalList.Count;
                            //db = GlobalFunc.ConnectDBByAdmin;
                            foreach (var approvaluri in approvalList)
                            {
                                //db.GetConsignmentApprover(Convert.ToInt32(approvaluri)); 
                                var approval = new HPSDK.ConsignmentApprover(database, approvaluri);
                                JW.WriteStartObject();

                                JW.WriteMember("Approver");
                                JW.WriteString(approval.Approver.FormattedName);
                                JW.WriteMember("Record");
                                JW.WriteString(approval.RecTitle);
                                JW.WriteMember("Consignment");
                                JW.WriteString(approval.Consignment.Number);
                                JW.WriteMember("Status");
                                JW.WriteString(approval.Status.ToString());
                                JW.WriteEndObject();
                            }
                        }
                        else
                        {
                            var ReviewRec = new HPSDK.Record(database, Convert.ToInt32(consignmentNo));

                            string searchstring = $"ConsignmentReviewUri:{consignmentNo}";
                            if (Pending != "")
                                searchstring += " AND ConsignmentReviewStatus:\"\"";
                            var reviewList = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchstring);
                            foreach (var recUri in reviewList)
                            {
                                //db.GetConsignmentApprover(Convert.ToInt32(approvaluri)); 
                                var rec = new HPSDK.Record(database, recUri);
                                JW.WriteStartObject();

                                JW.WriteMember("Approver");
                                JW.WriteString(ReviewRec.OwnerLocation.FormattedName);
                                JW.WriteMember("Record");
                                JW.WriteString(rec.Title);
                                JW.WriteMember("Consignment");
                                JW.WriteString(ReviewRec.Title);

                                string RecStatus = rec.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Review Status")).AsString();
                                JW.WriteMember("Status");
                                JW.WriteString(RecStatus == "" ? "Pending Approval" : RecStatus);
                                JW.WriteEndObject();
                            }
                        }
                    }

                    JW.WriteEndArray();

                    JW.WriteMember("totalrows");
                    JW.WriteNumber(count);

                    JW.WriteEndObject();
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString("");
                    JW.WriteEndObject();
                }
                GlobalFunc.Log(ex);
            }
        }
        
        private void SaveConsignmentHPE()
        {
            string cdtType = Request["cdtType"];
            if (cdtType == "" || cdtType == null)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString("The type parameter error of Consignment.");
                    JW.WriteEndObject();
                }
            }
            else
            {
                string number = Request["txtNumber"];
                string descr = Request["txtDescription"];
                string archivist = Request["txtResponsibleArchivist"];
                string cutOffDate = Request["txtCutOffDate"];
                bool isConversion = false; //Convert.ToBoolean(Request["isConversion"]);
                string method = "";
                string consignmentUriFlag = "";
                HPSDK.Consignment consig = null;

                try
                {
                    HPSDK.ConsignmentDisposalType type;
                    switch (cdtType)
                    {
                        case "New Review Consignment":
                            method = "Review";
                            type = HPSDK.ConsignmentDisposalType.Archive;
                            break;
                        case "New Archive Consignment":
                            method = "Archive";
                            type = HPSDK.ConsignmentDisposalType.Archive;
                            break;
                        case "New Destroy Consignment":
                            method = "Destroy";
                            type = HPSDK.ConsignmentDisposalType.Destroy;
                            break;
                        case "New Transfer Consignment":
                            method = "Transfer";
                            type = HPSDK.ConsignmentDisposalType.Transfer;
                            break;
                        default:
                            type = HPSDK.ConsignmentDisposalType.Archive;
                            break;
                    }
                    using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDB()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
                    {
                        if (!CheckDuplicateConsignmentNumber(database, number))
                        {
                            using (JsonWriter JW = CreateJsonWriter(Response.Output))
                            {
                                JW.WriteStartObject();
                                JW.WriteMember("success");
                                JW.WriteBoolean(false);
                                JW.WriteMember("message");
                                JW.WriteString($"Consignment with 'Number' of '{number}' already exists.");
                                JW.WriteEndObject();
                            }
                            return;
                        }
                        HPSDK.Location loc = new HPSDK.Location(database, archivist); // db.GetLocation(archivist);
                        if (!loc.HasPermission(HPSDK.UserPermissions.RecordArchivist))
                        {
                            using(JsonWriter JW = CreateJsonWriter(Response.Output))
                            {
                                JW.WriteStartObject();
                                JW.WriteMember("success");
                                JW.WriteBoolean(false);
                                JW.WriteMember("message");
                                JW.WriteString($"{archivist} needs the permission \"Record Archivist\" to be the responsible archivist on this consignment.");
                                JW.WriteEndObject();
                            }
                            return;
                        }
                        if (method != "Review")
                        {
                            var trimMainObjectSearch = new HPSDK.TrimMainObjectSearch(database, HPSDK.BaseObjectTypes.Record);
                            //only get part records into consignment
                            trimMainObjectSearch.SetSearchString("ConsignmentUriFlag:\"\" and not closedOn:blank");  //("not schedule[IsReview] and not closedOn:blank"); // not IsInherit");
                            //trimMainObjectSearch.SetSearchString("not schedule[IsReview]"); 

                            //consig = db.NewConsignment(type, DateTime.Parse(cutOffDate), "", "", true, true, true, true, true);
                            consig = new HPSDK.Consignment(database, type, DateTime.Parse(cutOffDate), trimMainObjectSearch, true, true, false, true, true);
                            consig.Number = number;
                            consig.Description = descr;
                            consig.SetFieldValue(new HPSDK.FieldDefinition(database, "ConsignmentMethod"), new HPSDK.UserFieldValue(method));
                            consig.SetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Created"), new HPSDK.UserFieldValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                            
                            //HPSDK.Location loc = new HPSDK.Location(database, archivist); // db.GetLocation(archivist);
                            consig.Archivist = loc;

                            if (type == HPSDK.ConsignmentDisposalType.Transfer)
                            {
                                consig.TransferLocation = new HPSDK.Location(database, archivist); //db.GetLocation(archivist);
                            }

                            consig.Save();
                            consignmentUriFlag = consig.Uri.ToString();
                            if (consig.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString() == "")
                            {
                                String hUID = GlobalFunc.DATASETID + "I" + consig.Uri;
                                consig.SetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri"), new HPSDK.UserFieldValue(hUID));
                                consig.Save();
                            }
                            var conversionList = Search(database, HPSDK.BaseObjectTypes.Record, $"consignment:{consig.Uri.ToString()} and schedule[IsConversion]");
                            if (conversionList.Count > 0)
                                isConversion = true;
                            consig.SetFieldValue(new HPSDK.FieldDefinition(database, "Is Conversion"), new HPSDK.UserFieldValue(isConversion));
                            consig.UpdateComment = "["+ method + " Consignment Added] " + number + " | Added | " + GlobalFunc.DATASETID + "I" + consig.Uri;
                            consig.Save();

                            AutoResolveEnclosedItem(consig.Uri.ToString());
                            //send mail complete consignment
                            if (loc != null)
                            {
                                var emailTemplate = new HPSDK.Record(database, _sendMailToNewConsignment);
                                SendMail(database, number, loc.EmailAddress, emailTemplate, indexUrl);
                            }

                        }
                        else
                        {
                            var reviewRec = new HPSDK.Record(new HPSDK.RecordType(database, "Review Consignment"));
                            reviewRec.Title = number;
                            reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "TitleField"), new HPSDK.UserFieldValue(number));
                            reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "ConsignmentMethod"), new HPSDK.UserFieldValue(method));

                            reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "Description"), new HPSDK.UserFieldValue(descr));
                            reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "Consignment CutOffDate"), new HPSDK.UserFieldValue(DateTime.Parse(cutOffDate)));
                            reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Status"), new HPSDK.UserFieldValue("New Consignment"));

                            reviewRec.SetOwnerLocation(loc);
                            reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Created"), new HPSDK.UserFieldValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                            //only get part records into consignment
                            string partsearch = "schedule[IsReview] and not hashold ConsignmentUriFlag:\"\" and ConsignmentReviewUri:\"\" and not closedOn:blank";// and not IsInherit";
                            using (var admindb = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                            {                                                                                                          //string partsearch = "schedule[IsReview] and not hashold and ConsignmentReviewUri:\"\"";// and not IsInherit";
                                var recList = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, partsearch);
                                reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Review Records Count"), new HPSDK.UserFieldValue(recList.Count));

                                reviewRec.Save();
                                string reviewUri = reviewRec.Uri.ToString();
                                consignmentUriFlag = reviewRec.Uri.ToString();
                                if (reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString() == "")
                                {
                                    String hUID = GlobalFunc.DATASETID + "I" + reviewUri;
                                    reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri"), new HPSDK.UserFieldValue(hUID));
                                    reviewRec.Save();
                                }


                                foreach (var recUri in recList)
                                {
                                    var rec = new HPSDK.Record(admindb, Convert.ToInt32(recUri));
                                    rec.SetFieldValue(new HPSDK.FieldDefinition(admindb, "Consignment Review Uri"), new HPSDK.UserFieldValue(reviewUri));
                                    rec.SetFieldValue(new HPSDK.FieldDefinition(admindb, "Consignment Review Approver"), new HPSDK.UserFieldValue(rec.OwnerLocation));
                                    rec.Save();
                                }

                                var conversionList = Search(database, HPSDK.BaseObjectTypes.Record, $"ConsignmentReviewUri:{reviewUri} and schedule[IsConversion]");
                                if (conversionList.Count > 0)
                                    isConversion = true;
                                reviewRec.SetFieldValue(new HPSDK.FieldDefinition(database, "Is Conversion"), new HPSDK.UserFieldValue(isConversion));
                                reviewRec.UpdateComment = "[Review Consignment Added] " + number + " | Added | " + GlobalFunc.DATASETID + "I" + reviewUri;
                                reviewRec.Save();
                                //send mail new consignment
                                if (loc != null)
                                {
                                    var emailTemplate = new HPSDK.Record(database, _sendMailToNewConsignment);
                                    SendMail(database, number, loc.EmailAddress, emailTemplate, indexUrl);
                                }
                            }
                        }
                        CheckingParentInConsignment(consignmentUriFlag, method);
                        using (JsonWriter JW = CreateJsonWriter(Response.Output))
                        {
                            JW.WriteStartObject();
                            JW.WriteMember("success");
                            JW.WriteBoolean(true);
                            JW.WriteMember("message");
                            JW.WriteString("Saved Consignments Successfully.");
                            JW.WriteEndObject();
                        }
                    }
                }
                catch (Exception ex)
                {
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(false);
                        JW.WriteMember("message");
                        JW.WriteString(ex.Message.ToString());
                        JW.WriteEndObject();
                    }
                }
                finally
                {
                    //GlobalFunc.ReleaseCOMObject(consig);
                    consig = null;
                }
            }
        }

        private bool CheckDuplicateConsignmentNumber(HPSDK.Database database, string number)
        {
            var checkreview = Search(database, HPSDK.BaseObjectTypes.Record, $"type[Review Consignment] and TitleField:\"{number}\"");
            if (checkreview.Count > 0)
                return false;

            var checklist = Search(database, HPSDK.BaseObjectTypes.Consignment, $"number:\"{number}\"");
            if (checklist.Count > 0)
                return false;

            return true;
        }

        private void AutoResolveEnclosedItem(string consiUri)
        {
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
            {
                var issues = SearchConsignmentIssue(database, consiUri);
                foreach (var iUri in issues)
                {
                    var issue = new HPSDK.ConsignmentIssue(database, iUri);
                    if (issue.Description.Contains("enclosed"))
                    {
                        try
                        {
                            var rec = new HPSDK.Record(database, issue.Record.Uri);
                            rec.IsEnclosed = true;
                            rec.Save();
                            issue.RemoveIfResolved();
                        }
                        catch
                        { }
                    }
                }
            }
        }

        private void CheckingParentInConsignment(string consignmentNo, string Method)
        {
            string searchstring = "";
            if (Method.ToLower() != "review")
                searchstring = $"consignment:{consignmentNo}";
            else
                searchstring = $"ConsignmentReviewUri:{consignmentNo}";
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
            {
                var list = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);

                foreach (var recUri in list)
                {
                    HPSDK.FieldDefinition fdFlag = new HPSDK.FieldDefinition(database, "Consignment Uri Flag");
                    HPSDK.UserFieldValue fvFlag = new HPSDK.UserFieldValue(consignmentNo);
                    UpdateConsignmentFlag(recUri, fdFlag, fvFlag, database);
                }

            }

        }

        private void UpdateConsignmentFlag(HPSDK.TrimURI recUri, HPSDK.FieldDefinition fdFlag, HPSDK.UserFieldValue fvFlag, HPSDK.Database database)
        {
            var rec = new HPSDK.Record(database, recUri);
            rec.SetFieldValue(fdFlag, fvFlag);
            rec.Save();
            //if ((int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.SubClass || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.Folder || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.SubFolder || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.Part)
            //{
            //    string searchstring = $"container:\"{rec.Container.Uri}\" and ConsignmentUriFlag:\"\" and not closedOn:blank";// and closedOn:blank";
            //    var recList = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchstring);
            //    //check rec schedule as same as container schedule
            //    if (recList.Count == 0 && rec.Container.RetentionSchedule != null && rec.RetentionSchedule.Uri == rec.Container.RetentionSchedule.Uri)
            //        UpdateConsignmentFlag(rec.Container.Uri, fdFlag, fvFlag, database);
            //}
        }

        //private void RemoveConsigmentFlag(string consignmentNo, HPSDK.Database database)
        //{
        //    HPSDK.FieldDefinition fdFlag = new HPSDK.FieldDefinition(database, "ConsignmentUriFlag");
        //    HPSDK.UserFieldValue fvFlag = new HPSDK.UserFieldValue("");

        //    string searchstring = $"ConsignmentUriFlag:\"{consignmentNo}\"";
        //    var recList = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchstring);
        //    foreach (var recUri in recList)
        //    {
        //        var rec = new HPSDK.Record(database, recUri);
        //        rec.SetFieldValue(fdFlag, fvFlag);
        //        rec.Save();
        //    }
        //}
        private void CompleteConsignmentGetListToRemove(string consignmentNo, HPSDK.Database database)
        {
            string searchstring = $"ConsignmentUriFlag:\"{consignmentNo}\"";
            var recList = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchstring);
            foreach (var recUri in recList)
            {
                RemoveConsignmentFlag(recUri, consignmentNo, database);
            }
        }
        private void RemoveConsignmentFlag(HPSDK.TrimURI recUri, string consignmentNo, HPSDK.Database database)
        {
            HPSDK.FieldDefinition fdFlag = new HPSDK.FieldDefinition(database, "Consignment Uri Flag");
            HPSDK.UserFieldValue fvFlag = new HPSDK.UserFieldValue("");
            var rec = new HPSDK.Record(database, recUri);
            if ((int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.Record || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.CompoundRecord)
                return; // rec = new HPSDK.Record(database, rec.Container.Uri);
            rec.SetFieldValue(fdFlag, fvFlag);
            rec.Save();
            //if ((int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.SubClass || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.Folder || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.SubFolder || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.Part)
            //{
            //    string searchstring = $"container:\"{rec.Container.Uri}\" and ConsignmentUriFlag:\"{consignmentNo}\"";
            //    var recList = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchstring);
            //    //check rec schedule as same as container schedule
            //    if (recList.Count == 0)
            //        RemoveConsignmentFlag(rec.Container.Uri, consignmentNo, database);
            //}
        }

        private void GetConsignment()
        {
            db = GlobalFunc.connectDB();
            string number = Request["Number"];

            Consignment co = db.GetConsignment(number);
            try
            {
                if (co != null)
                {
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(true);
                        JW.WriteMember("Number");
                        JW.WriteString(co.Number);
                        JW.WriteMember("Description");
                        JW.WriteString(co.Description);
                        JW.WriteMember("Archivist");
                        JW.WriteString(co.Archivist.SortName);
                        JW.WriteMember("CutoffDate");
                        JW.WriteString(co.CutoffDate.Date.ToString("yyyy-MM-dd"));
                        JW.WriteMember("isConversion");
                        JW.WriteBoolean(co.GetUserField(db.GetFieldDefinition("Is Conversion")));
                        JW.WriteEndObject();
                    }
                }
                else
                {
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(false);
                        JW.WriteMember("message");
                        JW.WriteString("The Consignment does not exist.");
                        JW.WriteEndObject();
                    }
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString(ex.Message.ToString());
                    JW.WriteEndObject();
                }
            }
            finally
            {
                GlobalFunc.ReleaseCOMObject(co);
                co = null;
            }
        }

        private void GetConsignmentHPE()
        {
            string number = Request["Number"];
            string uri = Request["Uri"];
            string method = Request["Method"];
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
            {
                try
                {
                    if (method.ToLower() != "review")
                    {
                        HPSDK.Consignment co = new HP.HPTRIM.SDK.Consignment(database, Convert.ToInt32(uri));
                        using (JsonWriter JW = CreateJsonWriter(Response.Output))
                        {
                            JW.WriteStartObject();
                            JW.WriteMember("success");
                            JW.WriteBoolean(true);
                            JW.WriteMember("Number");
                            JW.WriteString(co.Number);
                            JW.WriteMember("Description");
                            JW.WriteString(co.Description);
                            JW.WriteMember("Archivist");
                            JW.WriteString(co.Archivist.SortName);
                            JW.WriteMember("CutoffDate");
                            JW.WriteString(co.CutoffDate.ToDateTime().ToString("yyyy-MM-dd"));
                            //JW.WriteMember("isConversion");
                            //JW.WriteBoolean(co.GetFieldValue(db.GetFieldDefinition("Is Conversion")));
                            JW.WriteEndObject();
                        }
                    }
                    else
                    {
                        HPSDK.Record rec = new HP.HPTRIM.SDK.Record(database, Convert.ToInt32(uri));
                        using (JsonWriter JW = CreateJsonWriter(Response.Output))
                        {
                            JW.WriteStartObject();
                            JW.WriteMember("success");
                            JW.WriteBoolean(true);
                            JW.WriteMember("Number");
                            JW.WriteString(rec.Title);
                            JW.WriteMember("Description");
                            JW.WriteString(rec.GetFieldValue(new HPSDK.FieldDefinition(database, "Description")).AsString());
                            JW.WriteMember("Archivist");
                            JW.WriteString(rec.OwnerLocation.SortName);
                            JW.WriteMember("CutoffDate");
                            JW.WriteString(rec.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment CutOffDate")).AsDate().ToDateTime().ToString("yyyy-MM-dd"));
                            //JW.WriteMember("isConversion");
                            //JW.WriteBoolean(co.GetFieldValue(db.GetFieldDefinition("Is Conversion")));
                            JW.WriteEndObject();
                        }
                    }
                }
                catch (Exception ex)
                {
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(false);
                        JW.WriteMember("message");
                        JW.WriteString(ex.Message.ToString());
                        JW.WriteEndObject();
                    }
                }
                finally
                {
                }
            }
        }

        private void ExportIssues()
        {
            string jsonString = Request["jsonString"];
            string consignmentName = Request["consignmentName"] + "";
            //var node = Json2Xml(jsonString);

            string[] json = jsonString.Split('|');
            System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
            DataTable dt = new DataTable();
            dt.TableName = consignmentName + "Issues";
            foreach (var jsonData in json)
            {
                Hashtable A = js.Deserialize<Hashtable>(jsonData);

                if (!dt.Columns.Contains("Title"))
                    dt.Columns.Add("Title", typeof(string));
                if (!dt.Columns.Contains("Classification code"))
                    dt.Columns.Add("Classification code", typeof(string));
                if (!dt.Columns.Contains("Classification path"))
                    dt.Columns.Add("Classification path", typeof(string));
                if (!dt.Columns.Contains("System Identifie"))
                    dt.Columns.Add("System Identifie", typeof(string));
                if (!dt.Columns.Contains("Issue"))
                    dt.Columns.Add("Issue", typeof(string));

                DataRow dr = dt.NewRow();
                dr["Title"] = A["Title"].ToString();
                dr["Classification code"] = A["Classificationcode"].ToString();
                dr["Classification path"] = A["Classificationpath"].ToString();
                dr["System Identifie"] = A["SystemIdentifie"].ToString();
                dr["Issue"] = A["Description"].ToString();

                dt.Rows.Add(dr);
            }
            //string = ConvertDataTableToXML(dt);
            string tmpFileName = "Consignment[" + consignmentName + "]Issues.csv";
            string file = DataTable2Csv(dt, tmpFileName);
            using (JsonWriter JW = CreateJsonWriter(Response.Output))
            {
                JW.WriteStartObject();
                JW.WriteMember("success");
                JW.WriteBoolean(true);
                JW.WriteMember("filename");
                JW.WriteString(tmpFileName);
                JW.WriteEndObject();
            }
        }

        //rex add 
        private void ConsignmentDetailsExport()
        {
            try
            {
                string consignmentNo = Request["consignmentNo"];
                string uri = Request.Form.Get("Uri") + "";
                string Method = Request["method"];
                string consignmentName = Request["consignmentName"];
                bool isComplete = Convert.ToBoolean(Request["isComplete"]);
                string searchstring = "";

                //Excel default server physical path 
                string FilePath = System.Configuration.ConfigurationManager.AppSettings["SaffronWorkArea"] + "\\Export\\";
                string FilePath2 = System.Configuration.ConfigurationManager.AppSettings["SaffronWorkArea"] + "\\TempFiles\\";
                string fileName = "sample.xlsx";
                string tempFilePath = Path.Combine(FilePath2, fileName);
                string exportFileName = consignmentName + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                string exportFilePath = Path.Combine(FilePath, exportFileName);
                //Determine if there is a server physical path and do not create it
                if (!Directory.Exists(FilePath))
                    Directory.CreateDirectory(FilePath);
                //Determine whether the ExcelL template file exists
                FileInfo file = new FileInfo(tempFilePath);
                if (!file.Exists)
                    return;
                //Determine whether the exported Excel file exists
                FileInfo exportFile = new FileInfo(exportFilePath);
                if (exportFile.Exists)
                    exportFile.Delete();
                //if (!exportFile.Exists)
                //{
                    XSSFWorkbook exportWorkbook = new XSSFWorkbook();  //Create a new xlsx workbook 
                    exportWorkbook.CreateSheet("Sheet1");
                    FileStream exportSF = new FileStream(exportFilePath, FileMode.Create);
                    exportWorkbook.Write(exportSF);
                    exportSF.Close();
                    exportWorkbook.Close();
                //}

                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDB())
                {
                    if (Method.ToLower() == "review")
                        searchstring = $"ConsignmentReviewUri:{consignmentNo}";
                    else if (isComplete && (Method.ToLower() == "destroy" || Method.ToLower() == "transfer"))
                        searchstring = $"ConsignmentUriFlag:{consignmentNo} and type:stub";
                    else
                        searchstring = $"consignment:{consignmentNo}";
                    int copycount = 2;
                    int rowIndex =0;
                    int curPage = 1;
                    var list = Search(database, HPSDK.BaseObjectTypes.Record, searchstring, "Classification code");
                    foreach (var recUri in list)
                    {
                        var rec = new HPSDK.Record(database, recUri);

                        string ContainerConsignment = "";
                        if (Method.ToLower() == "review")
                        {
                            ContainerConsignment = GetFieldValue(database, "ConsignmentReviewUri", rec.Container);
                        }
                        else if (isComplete && (Method.ToLower() == "destroy" || Method.ToLower() == "transfer"))
                        {
                            ContainerConsignment = rec.Container.RecordType.Name == "Stub" ? GetFieldValue(database, "ConsignmentUriFlag", rec.Container) : (rec.Container.ConsignmentObject == null ? "" : rec.Container.ConsignmentObject.Uri.ToString());
                        }
                        else
                        {
                            ContainerConsignment = rec.Container.ConsignmentObject == null ? "" : rec.Container.ConsignmentObject.Uri.ToString();
                        }
                        if (ContainerConsignment == consignmentNo)
                            continue;

                        //Exec export data operation
                        InputToExcel(database,rec,ref copycount,ref curPage,ref rowIndex, tempFilePath, exportFilePath, Method);
                        //Recursive loop
                        GetAggregationToReport(database, consignmentNo, rec.Uri.ToString(), Method,ref copycount, ref curPage,ref rowIndex, tempFilePath, exportFilePath, isComplete);
                    }
                    
                    FileStream exportSF2 = new FileStream(exportFilePath, FileMode.Open, FileAccess.Read);
                    XSSFWorkbook exportWorkbook2 = new XSSFWorkbook(exportSF2);
                    ISheet exportWorksheet = exportWorkbook2.GetSheet("Sheet1") as ISheet;//Get the exported sheet
                    IRow row = exportWorksheet.GetRow(0);
                    ICellStyle cellStyle = null;
                    for (int i = 0; i <= exportWorksheet.LastRowNum; i++)
                    {
                        row = exportWorksheet.GetRow(i);
                        if (row != null)
                        {
                            //Change the total number of pages
                            string value = row.GetCell(5).ToString();
                            if (value.Contains("TotalPage"))
                            {
                                value = value.Replace("TotalPage", "" + (curPage - 1) + "");
                                row.GetCell(5).SetCellValue(value);
                                row.GetCell(5).CellStyle = Getcellstyle(exportWorkbook2, cellStyle, row.GetCell(5), 1);
                            }
                        }
                    }

                    //Remove the last time there was an extra template line
                    if (rowIndex < (curPage-1) * 18)
                    {
                        IEnumerator cells = null;
                        ICell cell = null;
                        for (int i = exportWorksheet.LastRowNum; i>=rowIndex; i--)
                        {
                            row = exportWorksheet.GetRow(i);
                            //exportWorksheet.RemoveRow(row);
                            //exportWorksheet.ShiftRows(i, i+ 1, -1);
                            cells = row.GetEnumerator();
                            while (cells.MoveNext())
                            {
                                cell = cells.Current as XSSFCell;
                                cell.SetCellValue("");
                                ICellStyle cellStyle_ = exportWorkbook2.CreateCellStyle();
                                cell.CellStyle = cellStyle_;
                            }
                            row.Height = 300;
                            //goto cccc;
                            //exportWorksheet.ShiftRows(i, i+ 1, -1);
                        }
                        //exportWorksheet.RemoveMergedRegion(exportWorksheet.LastRowNum);
                        //exportWorksheet.ShiftRows(rowIndex + 1, rowIndex + 8, -8);
                    }
                    //Save excel
                    FileStream filess = File.OpenWrite(exportFilePath);
                    exportWorkbook2.Write(filess);
                    filess.Close();
                    exportWorkbook2.Close();
                }
                
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(true);
                    JW.WriteMember("filename");
                    JW.WriteString(exportFileName);
                    JW.WriteEndObject();
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember(ex.Message.ToString());
                    JW.WriteBoolean(true);
                    JW.WriteEndObject();
                }
            }
        }

        private void InputToExcel(HPSDK.Database database, HPSDK.Record rec, ref int copycount, ref int curPage, ref int rowIndex, string tempFilePath, string exportFilePath,string method)
        {

            FileStream temFileRead = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook temWorkbook = new XSSFWorkbook(temFileRead);
            FileStream exportSF2 = new FileStream(exportFilePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook exportWorkbook2 = new XSSFWorkbook(exportSF2);
            ISheet temWorksheet = temWorkbook.GetSheet("Sheet1") as ISheet;//Get a template sheet
            ISheet exportWorksheet = exportWorkbook2.GetSheet("Sheet1") as ISheet;//Get the exported sheet
            IRow temRow = null;
            ICell temCell = null;
            IRow expRow = null;
            ICell expCell = null;
            try
            {
                //Do you need to re-create excel typesetting
                if (copycount % 2 == 0)
                {
                    //Create excel layout
                    GetExprotExcel(exportWorkbook2, temWorksheet, exportWorksheet, temRow, temCell, expRow, expCell, copycount, rowIndex);
                   
                    //The first line
                    //Report
                    exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue("DisposalReport");
                    //Date
                    exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(DateTime.Now.ToString("MM.dd.yyyy\rHH:mm:ss"));
                    rowIndex++;
                    
                    //The second line
                    //User
                    exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(GlobalFunc.UserName.ToString());
                    //page
                   exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue("Page " + curPage + " of " + "TotalPage");
                    curPage++;
                    rowIndex++;
                }
                copycount++;
                //Third line
                //Title
                if (rec.Title != null)
                    exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(rec.Title.ToString());
                //Classification Code
                exportWorksheet.GetRow(rowIndex).GetCell(3).SetCellValue(GetFieldValue(database, "Classification code", rec).ToString());
                //Date Closed
               
                if (rec.DateClosed != null)
                    exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(rec.DateClosed.Year != -1 ? rec.DateClosed.ToDateTime().ToString("MM.dd.yyyy") : "");
                rowIndex++;
                //Fourth line
                //Date Opened
                if (rec.DateCreated != null)
                    exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(rec.DateCreated.Year != -1 ? rec.DateCreated.ToDateTime().ToString("MM.dd.yyyy") : "");
                //Owner
                if (rec.OwnerLocation != null)
                    exportWorksheet.GetRow(rowIndex).GetCell(3).SetCellValue(rec.OwnerLocation.Name.ToString());
                //Relation - GRS Disposal Schedule Identifer"
                exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(GetFieldValue(database, "GRS disposal schedule identifier", rec).ToString());
                rowIndex++;
                ////Fifth line 
                //Part No.
                exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(GetFieldValue(database, "Part number", rec).ToString());
                //Security Classification
                if (rec.Security != null)
                    exportWorksheet.GetRow(rowIndex).GetCell(3).SetCellValue(rec.Security.ToString());
                //Security Classification Type
                if (rec.Security.ToString() != "CONFIDENTIAL" && rec.Security.ToString() != "UNCLASSIFIED")
                    exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(GetFieldValue(database, "Security Classification Type", rec).ToString());
                rowIndex++;
                ////The sixth line
                //Retention Period
                string retPeriod = "";
                string triggerType = "";
                int afterDays, afterMonths, afterYears = 0;
                string disDateuture = "";
                string eventTrigger = "";
                string isTrigge = "No";
                if (rec.RecordType.Name != "Record" && rec.RecordType.Name != "Compound Record" && rec.RecordType.Name != "Compound")
                    if (rec.RetentionSchedule != null)
                    {
                        retPeriod = "P";
                        for (int i = 0; i < rec.RetentionSchedule.ChildTriggers.Count; i++)
                        {
                            triggerType = rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).TriggerType.ToString();
                            afterDays = rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).AfterDays;
                            afterMonths = rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).AfterMonths;
                            afterYears = rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).AfterYears;
                            //if (rec.ConsignmentObject != null)
                            //if (rec.ConsignmentObject.DisposalMethod.ToString() == GetRecordDisp(triggerType))
                            if (method.ToString().ToLower() == GetRecordDisp(triggerType).ToLower())
                            {
                                if (rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).UserDefinedDate != null)
                                    if (rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).UserDefinedDate.Name == "Event Trigger Date")
                                        isTrigge = "Yes";
                                if (afterYears != 0)
                                    retPeriod += afterYears.ToString().PadLeft(2, '0') + "Y";
                                if (afterMonths != 0)
                                    retPeriod += afterMonths.ToString().PadLeft(2, '0') + "M";
                                if (afterDays != 0)
                                    retPeriod += afterDays.ToString().PadLeft(2, '0') + "D";
                                if (afterYears != 0 || afterMonths != 0 || afterDays != 0)
                                    retPeriod = retPeriod + ",P";
                                if (rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).FixedDate.Year != -1)
                                    disDateuture += rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).FixedDate.ToDateTime().ToString("MM.dd.yyyy") + ",";
                                eventTrigger += GetDateType(rec.RetentionSchedule.ChildTriggers.getItem(Convert.ToUInt32(i)).DateType.ToString()) + ",";
                            }
                        }
                        retPeriod = retPeriod != "P" ? retPeriod.Remove(retPeriod.LastIndexOf(","), 2) : "";
                        disDateuture = disDateuture != "" ? disDateuture.Remove(disDateuture.LastIndexOf(","), 1) : "";
                        eventTrigger = eventTrigger != "" ? eventTrigger.Remove(eventTrigger.LastIndexOf(","), 1) : "";
                    }
                exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(retPeriod);
                //Disposal Action
                if (rec.RecordType.Name != "Record" && rec.RecordType.Name != "Compound Record" && rec.RecordType.Name != "Compound" && rec.RecordType.Name!="Stub")
                {
                    string scheduleValue = rec.RetentionSchedule==null?"": rec.RetentionSchedule.GetFieldValue(new HPSDK.FieldDefinition(database, "Disposal Action")).ToString();
                    exportWorksheet.GetRow(rowIndex).GetCell(3).SetCellValue(scheduleValue);
                }
                //Disposal Date - Future
                exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(disDateuture);
                rowIndex++;

                //The seventh line
                //Event Trigger
                exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(eventTrigger);
                //RDS Title
                if (rec.RecordType.Name != "Record" && rec.RecordType.Name != "Compound Record" && rec.RecordType.Name != "Compound")
                {
                    if (rec.RetentionSchedule != null)
                        exportWorksheet.GetRow(rowIndex).GetCell(3).SetCellValue(rec.RetentionSchedule.Title.ToString());
                    //Folder Type new HPSDK.FieldDefinition(database, "Folder Type") 
                    exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(GetFieldValue(database, "Folder Type", rec).ToString());
                }
                rowIndex++;
                //The eighth line 
                //Event Trigger - External
                if (rec.RecordType.Name == "Record" || rec.RecordType.Name == "Compound Record" || rec.RecordType.Name == "Compound")
                    isTrigge = "";
                exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(isTrigge);
                //Disposition Date
                if (rec.DisposalDate != null)
                    exportWorksheet.GetRow(rowIndex).GetCell(3).SetCellValue(rec.DisposalDate.Year != -1 ? rec.DisposalDate.ToDateTime().ToString("MM.dd.yyyy") : "");
                //External Event Trigger Date
                if (rec.RecordType.Name != "Record" && rec.RecordType.Name != "Compound Record" && rec.RecordType.Name != "Compound")
                {
                     DateTime eventTriggerDate = rec.GetFieldValue(new HPSDK.FieldDefinition(database, "Event Trigger Date")).AsDate();
                     exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(eventTriggerDate.Year != 1 ? eventTriggerDate.ToString("MM.dd.yyyy") : "");
                }
                rowIndex++;
                //The ninth line  
                //Classification Path
                exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(GetFieldValue(database, "Classification Path", rec).ToString());
                //Responsible Officer
                exportWorksheet.GetRow(rowIndex).GetCell(5).SetCellValue(GetFieldValue(database, "Responsible Officer", rec).ToString());
                rowIndex++;
                //The tenth line
                //Remarks
                string remarks = rec.Notes.Replace("\r\n", "").ToString();
                exportWorksheet.GetRow(rowIndex).GetCell(1).SetCellValue(remarks);
                rowIndex++;
                //Save excel
                FileStream filess = File.OpenWrite(exportFilePath);
                exportWorkbook2.Write(filess);
                filess.Close();
                exportWorkbook2.Close();
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember(ex.Message.ToString());
                    JW.WriteBoolean(true);
                    JW.WriteEndObject();
                }
            }

        }

        private string GetDateType(string DateType)
        {
            string DisplayName = "";
            //SchTrigger sc = null;
            //sc.DateType = tgTriggerTypes.
            switch (DateType)
            {
                case "ContentDateClosed":
                    DisplayName = "Latest Date Closed of Contents";
                    break;
                case "ContentDateCreated":
                    DisplayName = "Latest Date Created of Contents";
                    break;
                case "ContentDateFinalized":
                    DisplayName = "Latest Date Declared As Final of Contents";
                    break;
                case "ContentDateLastAction":
                    DisplayName = "Latest Last Action Date of Contents";
                    break;
                case "ContentDateModified":
                    DisplayName = "Latest Date Modified of Contents";
                    break;
                case "DateArchived":
                    DisplayName = "Date Archived";
                    break;
                case "DateClosed":
                    DisplayName = "Date Closed";
                    break;
                case "DateCreated":
                    DisplayName = "Date Opened";
                    break;
                case "DateFinalized":
                    DisplayName = "Date Declared As Final";
                    break;
                case "DateInactive":
                    DisplayName = "Date Inactive";
                    break;
                case "DateLastAction":
                    DisplayName = "Last Action Date";
                    break;
                case "DateModified":
                    DisplayName = "Date Modified";
                    break;
                case "DateRegistered":
                    DisplayName = "Date Registered";
                    break;
                case "FixedDate":
                    DisplayName = "Fixed Date";
                    break;
                case "LastPartDateClosed":
                    DisplayName = "Latest Part Date Closed";
                    break;
                case "LastPartDateCreated":
                    DisplayName = "Latest Part Date Created";
                    break;
                case "RootDateCreated":
                    DisplayName = "First Part Date Created";
                    break;
                case "UserDefined":
                    DisplayName = "Event Trigger Date"; //"User Defined";
                    break;
                default:
                    DisplayName = "Date Created";
                    break;
            }
            return DisplayName;
        }

        private string GetRecordDisp(string RecordDisp)
        {
            string DisplayName = "";
            //SchTrigger sc = null;
            //sc.TriggerType = rdRecordDisp.
            switch (RecordDisp)
            {
                case "Destroyed":
                    DisplayName = "Destroy";
                    break;
                case "Active":
                    DisplayName = "Active";
                    break;
                case "Transferred":
                    DisplayName = "Transfer";
                    break;
                case "ArchivedInterim":
                    DisplayName = "Interim Archive";
                    break;
                case "ArchivedLocal":
                    DisplayName = "Archive";
                    break;
                case "Inactive":
                    DisplayName = "Review";
                    break;
                default:
                    DisplayName = "Archive";
                    break;
            }
            return DisplayName;
        }

        private static ICellStyle Getcellstyle(IWorkbook exportWorkbook2, ICellStyle cellStyle, ICell expCell, int row)
        {
            //Create a style object  
            cellStyle = exportWorkbook2.CreateCellStyle();
            //Set vertical alignment 
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            //Set automatic wrap 
            cellStyle.WrapText = true;
            IFont expFont = exportWorkbook2.CreateFont();
            expFont.FontHeightInPoints = 9;
            expFont.FontName = "Yu Mincho";
            //if (expCell.RowIndex == 0 || expCell.RowIndex == 1)
            if (row == 0 || row == 1)
            {
                if (expCell.ColumnIndex == 2)
                {
                    cellStyle.Alignment = HorizontalAlignment.Center;
                }
                if (expCell.ColumnIndex == 5)
                {
                    cellStyle.Alignment = HorizontalAlignment.Right;
                }
            }
            if (expCell.RowIndex % 18 == 1 || row == 9)
            {
                cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
            }
            //if (expCell.ColumnIndex == 0 || expCell.ColumnIndex == 2 || (expCell.ColumnIndex == 4&& 
            //    (expCell.RowIndex+2)%8!=0 && expCell.RowIndex != 6 && (expCell.RowIndex+2)%16!= 0 && expCell.RowIndex != 14) ||
            //     ((expCell.RowIndex  == 8|| expCell.RowIndex == 16) && expCell.ColumnIndex == 3) ||
            //     ((expCell.RowIndex + 2) % 8 == 0 && expCell.ColumnIndex == 3&& expCell.RowIndex!=6) ||
            //     ((expCell.RowIndex + 2) % 16 == 0 && expCell.ColumnIndex == 3))
            if (row != 0 && row != 1)
                if (expCell.ColumnIndex == 0 || expCell.ColumnIndex == 4 || expCell.ColumnIndex == 2)
                {
                    expFont.Boldweight = (short)FontBoldWeight.Bold;
                }
            if (row == 1&& expCell.ColumnIndex==5)
            {
                expFont.Boldweight = (short)FontBoldWeight.Bold;
            }
            if ((row == 8||row==16) && expCell.ColumnIndex == 2)
            {
                expFont.Boldweight = (short)FontBoldWeight.None;
            }

            cellStyle.SetFont(expFont);
            return cellStyle;
        }

        private static void GetExprotExcel(IWorkbook exportWorkbook2, ISheet temWorksheet, ISheet exportWorksheet, IRow temRow, ICell temCell, IRow expRow, ICell expCell, int refcount, int rowIndex)
        {
            System.Collections.IEnumerator cells = null;
            ICellStyle cellStyle = null;
            ICell cell = null;
            int cellWidth = 0;
            for (int row = 0; row <= 17; row++)
            {

                temRow = temWorksheet.GetRow(row);
                cells = temRow.GetEnumerator();
                expRow = exportWorksheet.CreateRow(row + rowIndex);
                expRow.Height = temRow.Height;
                //for (int col = 0; col <= 5; col++)
                //{
                //temCell = temRow.GetCell(col);
                //expCell = expRow.CreateCell(col);
                while (cells.MoveNext())
                {
                    cell = cells.Current as XSSFCell;
                    cellWidth = temWorksheet.GetColumnWidth(cell.ColumnIndex);
                    exportWorksheet.SetColumnWidth(cell.ColumnIndex, cellWidth);
                    temCell = temRow.GetCell(cell.ColumnIndex);
                    expCell = expRow.CreateCell(cell.ColumnIndex);
                    expCell.SetCellValue(temCell.StringCellValue);
                    //Set cell style
                    expCell.CellStyle = Getcellstyle(exportWorkbook2, cellStyle, expCell,row);
                }
                //}
                if (row == 0|| row == 1)
                {
                    exportWorksheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex+ row, rowIndex+ row, 2, 3));
                }
                if (row == 8 || row == 16)
                {
                    exportWorksheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex + row, rowIndex + row, 1, 3));
                }
                //if (rowIndex %18== 0 || rowIndex %19== 0||rowIndex==1)
                //{
                //    exportWorksheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 2, 3));
                //}
                if (row == 9|| row == 17)
                {
                    exportWorksheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex+ row, rowIndex + row, 1, 5));
                }
                //if (rowIndex %9== 0 && rowIndex != 0 )
                //{
                //    exportWorksheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(row, row, 1, 5));
                //}
            }
        }

        private void GetAggregationToReport(HPSDK.Database database, string consignmentNo,string uri,string Method,ref int copycount, ref int curPage, ref int rowIndex,string tempFilePath, string exportFilePath, bool isComplete)
        {
            try
            {
                var container = new HPSDK.Record(database, Convert.ToInt32(uri));
                string FaValuestring = "";
                if (container.Security.ToUpper().Equals("CONFIDENTIAL"))
                {
                    Saffron.User user = new Saffron.User();
                    FaValuestring = user.CheckTFADetails(Session.SessionID).ToUpper();
                }
                if (FaValuestring == "")
                { 
                    return;
                }
                string searchstring = $"container:\"{uri}\""; // and not closedOn:blank";
                var list = Search(database, HPSDK.BaseObjectTypes.Record, searchstring, "Classification code");

                foreach (var recUri in list)
                {
                    var rec = new HPSDK.Record(database, recUri);

                    if (rec.RecordType.Name != "Record" && rec.RecordType.Name != "Compound Record")
                    {
                        string Consignment = "";
                        if (Method.ToLower() == "review")
                        {
                            Consignment = GetFieldValue(database, "ConsignmentReviewUri", rec);
                        }
                        else if (isComplete && (Method.ToLower() == "destroy" || Method.ToLower() == "transfer"))
                        {
                            Consignment = rec.RecordType.Name == "Stub" ? GetFieldValue(database, "ConsignmentUriFlag", rec) : (rec.ConsignmentObject == null ? "" : rec.ConsignmentObject.Uri.ToString());
                        }
                        else
                        {
                            Consignment = rec.ConsignmentObject == null ? "" : rec.ConsignmentObject.Uri.ToString();
                        }
                        if (Consignment != consignmentNo)
                            continue;
                    }
                    InputToExcel(database,rec,ref copycount, ref curPage, ref rowIndex, tempFilePath, exportFilePath,Method);
                    GetAggregationToReport(database, consignmentNo, rec.Uri.ToString(), Method,ref copycount, ref curPage, ref rowIndex, tempFilePath, exportFilePath, isComplete);
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(true);
                    JW.WriteEndObject();
                }
            }
            
        }

        private void DownloadFile()
        {
            string fileName = Request["file"];
            string ServerPath = System.Configuration.ConfigurationManager.AppSettings["SaffronWorkArea"] + "\\Export";
            System.IO.Stream iStream = null;

            // Buffer to read 10K bytes in chunk:
            byte[] buffer = new Byte[10000];

            // Length of the file:
            int length;

            // Total bytes to read:
            long dataToRead;

            string filePath = Path.Combine(ServerPath, fileName);
            //URL = filePath;
            System.IO.FileInfo TargetFile = new System.IO.FileInfo(filePath);
            TargetFile.IsReadOnly = false;
            Response.Clear();


            try
            {
                // Open the file.
                iStream = new System.IO.FileStream(filePath, System.IO.FileMode.Open,
                            System.IO.FileAccess.Read, System.IO.FileShare.Read);

                // Total bytes to read:
                dataToRead = iStream.Length;


                Response.ContentType = "application/octet-stream";
                Response.AddHeader("Content-Disposition", "attachment; filename*=UTF-8''" + Uri.EscapeDataString(fileName));

                // Read the bytes.
                while (dataToRead > 0)
                {
                    // Verify that the client is connected.
                    if (Response.IsClientConnected)
                    {
                        // Read the data in buffer.
                        length = iStream.Read(buffer, 0, 10000);

                        // Write the data to the current output stream.
                        Response.OutputStream.Write(buffer, 0, length);

                        // Flush the data to the HTML output.
                        Response.Flush();

                        buffer = new Byte[10000];
                        dataToRead = dataToRead - length;
                    }
                    else
                    {
                        //prevent infinite loop if user disconnects
                        dataToRead = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                // Trap the error, if any.
                Response.Write("Error : " + ex.Message);
            }
            finally
            {
                if (iStream != null)
                {
                    //Close the file.
                    iStream.Close();
                }

                //DirectoryInfo dirInfo = new DirectoryInfo(ServerPath);

                ////Delete the files 
                //foreach (FileInfo file in dirInfo.GetFiles())
                //{
                //    file.Delete();
                //}
            }
        }
    
        //private string ConvertDataTableToXML(DataTable xmlDS)
        //{
        //    MemoryStream stream = null;
        //    XmlTextWriter writer = null;
        //    try
        //    {
        //        stream = new MemoryStream();
        //        writer = new XmlTextWriter(stream, Encoding.Default);
        //        xmlDS.WriteXml(writer);
        //        int count = (int)stream.Length;
        //        byte[] arr = new byte[count];
        //        stream.Seek(0, SeekOrigin.Begin);
        //        stream.Read(arr, 0, count);
        //        UTF8Encoding utf = new UTF8Encoding();
        //        return utf.GetString(arr).Trim();
        //    }
        //    catch(Exception ex)
        //    {
        //        return String.Empty;
        //    }
        //    finally
        //    {
        //        if (writer != null) writer.Close();
        //    }
        //}

        private string DataTable2Csv(DataTable dt, string filename)
        {
            string EXPORT_DIR = System.Configuration.ConfigurationManager.AppSettings["SaffronWorkArea"] + "\\Export";
            string filePath = Path.Combine(EXPORT_DIR, filename);
            if (!Directory.Exists(EXPORT_DIR))
                Directory.CreateDirectory(EXPORT_DIR);

            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
                file.Delete();

            string csvString = "";
            foreach (DataColumn dc in dt.Columns)
            {
                csvString += dc.ColumnName + ",";
            }
            WriteCsv(csvString.TrimEnd(','), filePath, true);

            foreach (DataRow dr in dt.Rows)
            {
                csvString = "";
                foreach (DataColumn dc in dt.Columns)
                {
                    csvString += "\"" + dr[dc].ToString() + "\",";
                }
                WriteCsv(csvString.TrimEnd(','), filePath,false);
            }
            return filePath;
        }

        private void WriteCsv(string csvString, string filePath, bool isCreate)
        {
            FileStream fs;
            StreamWriter sw;

            if (isCreate)
            {
                fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            }
            else
            {
                fs = new FileStream(filePath, FileMode.Append, FileAccess.Write);
            }


            sw = new StreamWriter(fs);
            sw.WriteLine(csvString);
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// Mark Description in Consignment
        /// </summary>
        /// <param name="database"></param>
        /// <param name="consignmentNo">consignemt/review consignment uri</param>
        /// <param name="recuri">aggregation uri</param>
        /// <param name="description"></param>
        /// <param name="method"></param>
        private void MarkDescriptionByRemove(HPSDK.Database database,string consignmentNo, string recuri, string description, string method, DateTime? date = null)
        {
            var fdRemoveText = new HPSDK.FieldDefinition(database, "Consignment Remove Text");
            string dateStr = "\"date\": \"" + (date == null ? "" : date.ToString()) + "\"" ;
            string removetext = "{ \"recUri\": "+ recuri + ", " + dateStr + ", \"description\": \""+ description + "\" }";
            string fieldtext = "";
            if (method.ToLower() != "review")
            {
                var consignment = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));
                fieldtext = consignment.GetFieldValue(fdRemoveText).AsString();
                if (fieldtext != "")
                    fieldtext += ",";
                fieldtext += removetext;
                consignment.SetFieldValue(fdRemoveText, new HPSDK.UserFieldValue(fieldtext));
                consignment.Save();
            }
            else
            {
                var consignment = new HPSDK.Record(database, Convert.ToInt32(consignmentNo));
                fieldtext = consignment.GetFieldValue(fdRemoveText).AsString();
                if (fieldtext != "")
                    fieldtext += ",";
                fieldtext += removetext;
                consignment.SetFieldValue(fdRemoveText, new HPSDK.UserFieldValue(fieldtext));
                consignment.Save();
            }
        }

        private void RemoveFromConsignment()
        {
            string ruris = Request["removeuri"];
            string isHoldRecord = "";
            try
            {
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
                {
                    string[] uriList = ruris.TrimEnd(',').Split(',');

                    foreach (var ruri in uriList)
                    {
                        var issue = new HPSDK.ConsignmentIssue(database, Convert.ToInt32(ruri));
                        //if (issue.Record.HasHold)
                        //{
                        //    isHoldRecord += $"[{GetFieldValue(database, "Classification code", issue.Record, "record")}],";
                        //    continue;
                        //}
                        var consignment = issue.Consignment;
                        string issuesDescription = issue.Description;
                        HPSDK.Record rec = null;
                        if (issue.Record.ConsignmentObject != null)
                        {
                            rec = issue.Record;
                            if (rec.ConsignmentObject.Uri != consignment.Uri)
                            {
                                var container = new HPSDK.Record(database, rec.Container.Uri);
                                issue.ResolveByRemovingRecordFromContainer();
                                rec.Refresh();
                                if (rec.Container == null) //if removed container, then add again
                                {
                                    rec.Container = container;
                                    rec.Save();
                                }
                            }
                            else
                                issue.ResolveByRemovingRecord();
                        }
                        else
                        {
                            rec = issue.Record;
                            //issue.ResolveByRemovingRecord();
                            consignment.RemoveRecord(rec);
                            issue.RemoveIfResolved();
                        }
                        string strHPEUri = rec.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                        rec.UpdateComment = "[Remove from consignment] Record remove from consignment | Removed |" + strHPEUri;
                        rec.Save();
                        if (issue.Record.ConsignmentObject != null && issue.Record.ConsignmentObject.Uri == consignment.Uri)
                        {
                            MarkDescriptionByRemove(database, consignment.Uri.ToString(), rec.Uri.ToString(), issuesDescription, "");

                            RemoveConsignmentFlag(rec.Uri, consignment.Uri.ToString(), database);
                        }
                    }
                    string message = "";
                    //if (isHoldRecord.Length > 0)
                    //{
                    //    message = isHoldRecord.TrimEnd(',') + " already on hold, can not remove from consignment!";
                    //}
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(true);
                        JW.WriteMember("msg");
                        JW.WriteString(message);
                        JW.WriteEndObject();
                    }
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("msg");
                    JW.WriteString(ex.Message.ToString());
                    JW.WriteEndObject();
                }
            }
        }

        private void ReturnToHome()
        {
            string uris = Request["uris"];
            bool hasError = false;
            string errormsg = "";
            string isHoldRecord = "";
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
            {
                string[] uriList = uris.TrimEnd(',').Split(',');
                foreach (var uri in uriList)
                {
                    if (uri == "")
                        continue;
                    try
                    {
                        var issue = new HPSDK.ConsignmentIssue(database, Convert.ToInt32(uri));
                        //var list = SearchIssuesAggregation(database, issue.Record.Name);
                        var rec = new HPSDK.Record(database, issue.Record.Uri);
                        //if (issue.Record.HasHold)
                        //{
                        //    isHoldRecord += $"[{GetFieldValue(database, "Classification code", issue.Record, "record")}],";
                        //    continue;
                        //}
                        rec.SetCurrentLocationAtHome();
                        string strHPEUri = rec.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                        rec.UpdateComment = "[Return to home] Record return to home | Returned to home |" + strHPEUri;
                        rec.Save();
                        issue.RemoveIfResolved();
                    }
                    catch (Exception ex)
                    {
                        hasError = true;
                        errormsg += ex.Message + "\r";
                    }
                }
                
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(!hasError);
                    JW.WriteMember("msg");
                    JW.WriteString(errormsg);
                    JW.WriteEndObject();
                }
            }
        }

        private void EncloseItem()
        {
            string uris = Request["uris"];
            bool hasError = false;
            string errormsg = "";
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
            {
                string[] uriList = uris.TrimEnd(',').Split(',');
                foreach (var uri in uriList)
                {
                    try
                    {
                        var issue = new HPSDK.ConsignmentIssue(database, Convert.ToInt32(uri));

                        var rec = new HPSDK.Record(database, issue.Record.Uri);
                        rec.IsEnclosed = true;
                        rec.Save();
                        issue.RemoveIfResolved();
                    }
                    catch (Exception ex)
                    {
                        hasError = true;
                        errormsg += ex.Message + "\r";
                    }
                }


                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(!hasError);
                    JW.WriteMember("msg");
                    JW.WriteString(errormsg);
                    JW.WriteEndObject();
                }
            }
        }

        private void ShowDetail()
        {
            try
            {
                
                    string consignmentNo = Request["consignmentNo"];
                    string uri = Request.Form.Get("Uri") + "";
                    string code = Request.Form.Get("Code") + "";
                    string rds = Request.Form.Get("RDS") + "";
                    string Method = Request["method"];
                    bool isComplete = Convert.ToBoolean(Request["isComplete"]);

                    string searchstring = "";
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDB()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
                {
                    if (uri != "source" && uri != "")
                    {
                        var container = new HPSDK.Record(database, Convert.ToInt32(uri));
                        string FaValuestring = "";
                        bool hasAccess = true;
                        if (container.Security.ToUpper().Equals("CONFIDENTIAL"))
                        {
                            Saffron.User user = new Saffron.User();
                            FaValuestring = user.CheckTFADetails(Session.SessionID).ToUpper();
                            if (FaValuestring == "0" || FaValuestring == "1")
                                hasAccess = true;
                            else
                                hasAccess = false;
                        }
                        if (hasAccess)
                        {
                            //Parent = new HPSDK.Schedule(database, Convert.ToInt32(uri));
                            searchstring = $"container:\"{uri}\""; // and not closedOn:blank";
                        }
                        else
                        {
                            using (JsonWriter JW = CreateJsonWriter(Response.Output))
                            {
                                JW.WriteStartObject();
                                JW.WriteMember("success");
                                JW.WriteBoolean(false);
                                JW.WriteMember("msg");
                                JW.WriteString("Confenditial data");
                                JW.WriteMember("data");
                                JW.WriteStartArray();
                                JW.WriteEndArray();

                                JW.WriteEndObject();
                            }
                            return;
                        }
                    }
                    else
                    {
                        if (Method.ToLower() == "review")
                            searchstring = $"ConsignmentReviewUri:{consignmentNo}";
                        else if (isComplete && (Method.ToLower() == "destroy" || Method.ToLower() == "transfer"))
                            searchstring = $"ConsignmentUriFlag:{consignmentNo}";
                        else
                            searchstring = $"consignment:{consignmentNo}";
                    }

                    //CheckingParentInConsignment(consignmentNo, Method);
                    var list = Search(database, HPSDK.BaseObjectTypes.Record, searchstring, "Classification code");  //SearchShowDetails(database, consignmentNo);
                    bool Leaf = false;

                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartArray();
                        foreach (var recUri in list)
                        {
                            var rec = new HPSDK.Record(database, recUri);
                            if (rec.RecordType.Name != "Record" && rec.RecordType.Name != "Compound Record")
                                if (uri == "source")
                                {
                                    string ContainerConsignment = "";
                                    if (Method.ToLower() == "review")
                                    {
                                        ContainerConsignment = GetFieldValue(database, "ConsignmentReviewUri", rec.Container);
                                    }
                                    else if (isComplete && (Method.ToLower() == "destroy" || Method.ToLower() == "transfer"))
                                    {
                                        ContainerConsignment = rec.Container.RecordType.Name == "Stub" ? GetFieldValue(database, "ConsignmentUriFlag", rec.Container) : (rec.Container.ConsignmentObject == null ? "" : rec.Container.ConsignmentObject.Uri.ToString());
                                    }
                                    else
                                    {
                                        ContainerConsignment = rec.Container == null ? "" : (rec.Container.ConsignmentObject == null ? "" : rec.Container.ConsignmentObject.Uri.ToString());
                                    }
                                    if (ContainerConsignment == consignmentNo)
                                        continue;
                                }
                                else
                                {
                                    string Consignment = "";
                                    if (Method.ToLower() == "review")
                                    {
                                        Consignment = GetFieldValue(database, "ConsignmentReviewUri", rec);
                                    }
                                    else if (isComplete && (Method.ToLower() == "destroy" || Method.ToLower() == "transfer"))
                                    {
                                        Consignment = rec.RecordType.Name == "Stub" ? GetFieldValue(database, "ConsignmentUriFlag", rec) : (rec.ConsignmentObject == null ? "" : rec.ConsignmentObject.Uri.ToString());
                                    }
                                    else
                                    {
                                        Consignment = rec.ConsignmentObject == null ? "" : rec.ConsignmentObject.Uri.ToString();
                                    }
                                    if (Consignment != consignmentNo)
                                        continue;
                                }
                            //HPSDK.FieldDefinition fdLocationHome = new HPSDK.FieldDefinition(database, "LocationHome");
                            //string value = rec.GetFieldValue(fdLocationHome).ToString();
                            if (rec.Contents != "")
                            {
                                Leaf = false;
                            }
                            else
                            {
                                Leaf = true;
                            }

                            JW.WriteStartObject();
                            JW.WriteMember("Uri");
                            JW.WriteString(rec.Uri.ToString());
                            JW.WriteMember("Classificationcode");
                            JW.WriteString(GetFieldValue(database, "Classification code", rec)); // rec.GetFieldValue(new HPSDK.FieldDefinition(database, "Classification code")).ToString());
                            JW.WriteMember("Classificationpath");
                            JW.WriteString(GetFieldValue(database, "Classification path", rec)); //rec.GetFieldValue(new HPSDK.FieldDefinition(database, "Classificationpath")).ToString());
                            JW.WriteMember("Title");
                            JW.WriteString(rec.Title);
                            JW.WriteMember("Owner");
                            JW.WriteString(rec.OwnerLocation.Name);
                            JW.WriteMember("Dateopen");
                            JW.WriteString(rec.DateCreated.Year == -1 ? "" : rec.DateCreated.ToString());
                            JW.WriteMember("Dateclosed");
                            JW.WriteString(rec.DateClosed.Year == -1 ? "" : rec.DateClosed.ToString());

                            string home = GetFieldValue(database, "Location - Home", rec, "folder");
                            int homeUri;
                            if (int.TryParse(home, out homeUri))
                            {
                                var location = new HPSDK.Location(database, homeUri);
                                if (location != null)
                                    home = location.Name;
                            }
                            JW.WriteMember("LocationHome");
                            JW.WriteString(home);
                            JW.WriteMember("LocationCurrent");
                            JW.WriteString(GetFieldValue(database, "Location - Current", rec, "folder"));
                            string Physical = "";
                            if (rec.RecordType.Name.ToLower() == "part" && !Leaf)
                            {
                                string searchPhysical = $"container:{rec.Uri}+ and not electronic and type:record";
                                Physical = Search(database, HPSDK.BaseObjectTypes.Record, searchPhysical).Count.ToString();
                            }
                            JW.WriteMember("recordcount");
                            JW.WriteString(Physical);

                            JW.WriteMember("RDS");
                            JW.WriteString(rec.RetentionSchedule == null ? rds : rec.RetentionSchedule.Name);
                            string disposalaction = rec.RetentionSchedule == null ? "" : rec.RetentionSchedule.GetFieldValue(new HPSDK.FieldDefinition(database, "Disposal Action")).AsString();
                            JW.WriteMember("Disposalaction");
                            JW.WriteString(disposalaction);

                            JW.WriteMember("number");
                            JW.WriteString(rec.Number);

                            string Icon = "";
                            if (rec.IsElectronic)
                            {
                                //Icon = "attachment";
                                Icon = GetIcon(ref rec, rec.Extension.ToLower());
                            }
                            else
                            {
                                Icon = GetRecordIcon(rec.RecordType);
                            }
                            JW.WriteMember("iconCls");
                            JW.WriteString(Icon);

                            String FolderTypeIcon = GetFolderIcon(rec, database);

                            JW.WriteMember("FolderTypeIcon");
                            JW.WriteString(FolderTypeIcon);

                            JW.WriteMember("SecurityClassification");
                            JW.WriteString(rec.Security.ToString());

                            JW.WriteMember("leaf");
                            JW.WriteBoolean(Leaf);

                            JW.WriteEndObject();
                        }
                        JW.WriteEndArray();
                    }
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(true);
                    JW.WriteEndObject();
                }
            }
        }

        private void DeleteConsignment()
        {
            string consignmentNo = Request["consignmentNo"];
            string method = Request["method"];
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDB())
            {
                try
                {
                    if (method.ToLower() == "review")
                    {
                        var reviewRec = new HPSDK.Record(database, Convert.ToInt32(consignmentNo));
                        reviewRec.UpdateComment = "Consignment Deleted - " + reviewRec.Title;
                        reviewRec.Delete();
                        var reList = Search(database, HPSDK.BaseObjectTypes.Record, $"consignmentReviewUri:{consignmentNo}");
                        using (var admindb = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                        {
                            HPSDK.FieldDefinition fdFlag = new HPSDK.FieldDefinition(admindb, "consignment Review Uri");
                            HPSDK.UserFieldValue fvFlag = new HPSDK.UserFieldValue("");
                            foreach (var reviewUri in reList)
                            {
                                var rec = new HPSDK.Record(database, reviewUri);
                                rec.SetFieldValue(fdFlag, fvFlag);
                                rec.Save();
                            }
                        }
                        string strHPEUri = reviewRec.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                        string reason = "Consignment Deleted.";
                        reviewRec.UpdateComment = " [Delete Consignment] " + reason + "| Deleted |" + strHPEUri;
                        reviewRec.Save();
                    }
                    else
                    {
                        var consi = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));                        
                        string strHPEUri = consi.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                        string reason = "Consignment Deleted";
                        consi.UpdateComment = " [Delete Consignment] " + reason + "| Deleted |" + strHPEUri;
                        consi.Delete();
                        consi.Save();
                    }
                    var dbadmin = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin();
                    CompleteConsignmentGetListToRemove(consignmentNo, dbadmin);
                    dbadmin.Disconnect();
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(true);
                        JW.WriteMember("message");
                        JW.WriteString("");
                        JW.WriteEndObject();
                    }
                }
                catch (Exception ex)
                {
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(false);
                        JW.WriteMember("message");
                        JW.WriteString(ex.Message);
                        JW.WriteEndObject();
                    }
                }
            }
        }

        private string GetFieldValue(HPSDK.Database database, string fieldName, HPSDK.Record rec, string strExclude = "")
        {
            string value = "";
            try
            {
                if (rec.RecordType.Name.ToLower() != strExclude.ToLower() || strExclude == "")
                {
                    value = rec.GetFieldValue(new HPSDK.FieldDefinition(database, fieldName)).ToString();
                }
            }
            catch
            {
                value = "";
            }
            return value;
        }

        private void GetSecuritAudit()
        {
            string conNo = Request["consignmentNo"];
            string conName = Request["consignmentName"];
            string method = Request["method"];
            try
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();

                    JW.WriteMember("results");
                    JW.WriteStartArray();
                    int count = 0;
                    int index = 0;
                    using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin()) //new HP.HPTRIM.SDK.Database { Id = GlobalFunc.DATASETID, WorkgroupServerName = workgroup })
                    {                    
                        var Audits = SearchConsignmentAudit(database, conNo, method);
                        count = Audits.Count;
                        foreach (var audituri in Audits)
                        {
                            var Audit = new HPSDK.History(database, audituri);
                            JW.WriteStartObject();
                            JW.WriteMember("event");
                            JW.WriteString(conName + " " + GetEventsDes(Audit));
                            JW.WriteMember("object");
                            JW.WriteString(conName + $"({conNo})");
                            JW.WriteMember("updateby");
                            JW.WriteString(Audit.LoginLocation.Name);
                            JW.WriteMember("date");
                            JW.WriteString(Audit.DoneOn.ToString());
                            JW.WriteMember("details");
                            JW.WriteString(Audit.EventDescription);
                            JW.WriteEndObject();
                        }
                        JW.WriteEndArray();

                        JW.WriteMember("totalrows");
                        JW.WriteNumber(count);

                        JW.WriteEndObject();
                        //database.Connect(); //Connect with either .ConnectAs or .Connect method

                        //Do something, replace example code here
                    }



                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("msg");
                    JW.WriteString(ex.Message.ToString());
                    JW.WriteEndObject();
                }
            }
        }

        private void CompleteReview()
        {
            string uri = Request["uri"];
            string method = Request["method"];
            try
            {
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                {
                    if (CheckCONFIDENTIAL(database, uri))
                    {
                        using (JsonWriter JW = CreateJsonWriter(Response.Output))
                        {
                            JW.WriteStartObject();
                            JW.WriteMember("success");
                            JW.WriteBoolean(false);
                            JW.WriteMember("message");
                            JW.WriteString("Consignment has confidential data, you do not have access to confidential data and cannot complete review this consignment.");
                            JW.WriteEndObject();
                        }
                        return;
                    }
                    if (method.ToLower() != "review")
                    {
                        var consi = new HPSDK.Consignment(database, Convert.ToInt32(uri));
                        if (consi != null)
                        {
                            consi.SetArchivistReviewComplete();
                            string hpeuri = consi.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).AsString();
                            consi.UpdateComment = "[Consignment Complete Review] " + consi.Number + " | Complete Review | " + hpeuri;
                            consi.Save();
                        }
                        else
                        {
                            using (JsonWriter JW = CreateJsonWriter(Response.Output))
                            {
                                JW.WriteStartObject();
                                JW.WriteMember("success");
                                JW.WriteBoolean(false);
                                JW.WriteMember("message");
                                JW.WriteString("The Consignment does not exist.");
                                JW.WriteEndObject();
                            }
                        }
                    }
                    else
                    {
                        var review = new HPSDK.Record(database, Convert.ToInt32(uri));
                        DateTime reviewDate = DateTime.Now;
                        review.SetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Review Date"), new HPSDK.UserFieldValue(reviewDate));
                        review.SetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Status"), new HPSDK.UserFieldValue("Date Review Complete:" + reviewDate.ToShortDateString() + " at " + reviewDate.ToShortTimeString()+";"));
                        string hpeuri = review.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).AsString();
                        review.UpdateComment = "[Consignment Complete Review] " + review.Title + " | Complete Review | " + hpeuri;
                        review.Save();
                    }
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(true);
                        JW.WriteMember("message");
                        JW.WriteString("Complete the Review of the Consignment.");
                        JW.WriteEndObject();
                    }
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString(ex.Message.ToString());
                    JW.WriteEndObject();
                }
            }
        }

        private void CheckOnHoldRecord()
        {
            string consignmentNo = Request["consignmentNo"];
            string Method = Request["Method"];
            string searchstring = "";
            string onholdRecord = "";
            if (Method.ToLower() != "review")
                searchstring = $"consignment:{consignmentNo}";
            else
                searchstring = $"ConsignmentReviewUri:{consignmentNo}";


            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
            {
                var list = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchstring);
                foreach (var recUri in list)
                {
                    var rec = new HPSDK.Record(database, recUri);
                    if (rec.ChildHolds.Count > 0)
                        onholdRecord += $"[{GetFieldValue(database, "Classification code", rec)}],";
                }
            }
            if (onholdRecord != "")
                onholdRecord = onholdRecord.TrimEnd(',') + " on hold, can not be complete.";

            using (JsonWriter JW = CreateJsonWriter(Response.Output))
            {
                JW.WriteStartObject();
                JW.WriteMember("success");
                JW.WriteBoolean(string.IsNullOrEmpty(onholdRecord));
                JW.WriteMember("msg");
                JW.WriteString(onholdRecord);
                JW.WriteEndObject();
            }
        }

        private void CompleteConsignment()
        {
            string consignmentNo = Request["consignmentNo"];
            string consignmentName = Request["consignmentName"];
            string dataElectronic = Request["dataElectronic"];
            string remark = Request["remark"];
            string Method = Request["Method"];
            string reason = Request["reason"] == null ? "" : Request["reason"];
            try
            {
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                {
                    if (CheckCONFIDENTIAL(database, consignmentNo))
                    {
                        using (JsonWriter JW = CreateJsonWriter(Response.Output))
                        {
                            JW.WriteStartObject();
                            JW.WriteMember("success");
                            JW.WriteBoolean(false);
                            JW.WriteMember("msg");
                            JW.WriteString("Consignment has confidential data, you do not have access to confidential data and cannot complete this consignment.");
                            JW.WriteEndObject();
                        }
                        return;
                    }
                    HPSDK.Record Review = null;
                    HPSDK.Consignment consignment = null;
                    if (Method != "Review")
                    {
                        CheckDiffConsignmentWithChild(database, consignmentNo);
                        consignment = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));
                    }
                    else
                        Review = new HPSDK.Record(database, Convert.ToInt32(consignmentNo));
                    //string Method = consignment.DisposalMethod.ToString();
                    //if(Method == "Destroy" || Method == "Transfer")
                    //    CheckDiffConsignmentWithChild(database, consignmentNo);
                    Report(database,consignmentNo, Method, consignment, Review, reason);
                    HPSDK.FieldDefinition fdRemark = new HPSDK.FieldDefinition(database, "Consignment Remark");
                    HPSDK.UserFieldValue fvRemark = new HPSDK.UserFieldValue(remark);
                    switch (Method)
                    {
                        case "Destroy":
                            UpdateElectronic(database, dataElectronic);
                            consignment.DoDisposal(false, new HPSDK.RecordType(database, "consignmentLog"), null);
                            consignment.SetFieldValue(fdRemark, fvRemark);
                            consignment.Save();
                            ModifyRecordType(consignmentNo, reason);
                            break;
                        case "Transfer":
                            UpdateElectronic(database, dataElectronic);
                            consignment.PrepareForTransfer(new HPSDK.RecordType(database, "consignmentLog"), null);
                            consignment.CompleteTransfer(false, DateTime.Now);
                            consignment.SetFieldValue(fdRemark, fvRemark);
                            consignment.Save();
                            ModifyRecordType(consignmentNo, reason);
                            break;
                        case "Archive":
                            consignment.DoDisposal(false, new HPSDK.RecordType(database, "consignmentLog"), null);
                            consignment.SetFieldValue(fdRemark, fvRemark);
                            consignment.Save();
                            //CompleteConsignmentGetListToRemove(consignmentNo, database);
                            break;
                        case "Review":
                            DateTime date = DateTime.Now;
                            string State = "Complete Consignment Date:" + date.ToShortDateString() + " at " + date.ToShortTimeString();
                            Review.SetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Status"), new HPSDK.UserFieldValue(State));
                            Review.SetFieldValue(new HPSDK.FieldDefinition(database, "Consignment IsComplete"), new HPSDK.UserFieldValue(true));
                            Review.SetFieldValue(fdRemark, fvRemark);
                            Review.Save();
                            CompleteConsignmentGetListToRemove(consignmentNo, database);
                            break;
                    }

                    if (Method != "Review")
                    {
                        string strHPEUri = consignment.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                        consignment.UpdateComment = "[Complete Consignment] Consignment Completed |Completed| " + strHPEUri;
                        consignment.Save();
                    }
                    else
                    {
                        string strHPEUri = Review.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                        Review.UpdateComment = "[Complete Consignment] Consignment Completed |Completed| " + strHPEUri;
                        Review.Save();
                    }
                    
                    
                    using (JsonWriter JW = CreateJsonWriter(Response.Output))
                    {
                        JW.WriteStartObject();
                        JW.WriteMember("success");
                        JW.WriteBoolean(true);
                        JW.WriteMember("msg");
                        JW.WriteString("");
                        JW.WriteEndObject();
                    }
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("msg");
                    JW.WriteString(ex.Message.ToString());
                    JW.WriteEndObject();
                }
            }
        }

        private bool CheckCONFIDENTIAL(HPSDK.Database database, string consignmentNo)
        {
            string searchstring = $"consignment:{consignmentNo} and securityLevel:[CONFIDENTIAL]";
            var list = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.Record, searchstring);
            string FaValuestring = "";
            if (list.Count > 0)
            {
                Saffron.User user = new Saffron.User();
                FaValuestring = user.CheckTFADetails(Session.SessionID).ToUpper();
                if (FaValuestring == "")
                {
                    return true;
                }
            }

            return false;
        }

        private void CheckDiffConsignmentWithChild(HPSDK.Database database, string consignmentNo)
        {
            string searchstring = $"consignment:{consignmentNo} and not container:[consignment:[{consignmentNo}]]";
            var toplist = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);

            checkRecConsignment(database, toplist, consignmentNo);
        }

        private void checkRecConsignment(HPSDK.Database database, HPSDK.TrimURIList list, string consignmentNo)
        {
            foreach (var recuri in list)
            {
                //string consignmentString = $"uri:{recuri}+ and consignment:{consignmentNo} and type[\"sub-class\",\"class\",\"folder\",\"sub-folder\",\"part\"]";
                //string allrecString = $"uri:{recuri}+ and type[\"sub-class\",\"class\",\"folder\",\"sub-folder\",\"part\"]";
                string allrecString = $"uri:{recuri}+ and not consignment:{consignmentNo} and (consignment[disposalDate:blank] or consignment[]) and type[\"sub-class\",\"class\",\"folder\",\"sub-folder\",\"part\"]";

                //var conslist = Search(database, HPSDK.BaseObjectTypes.Record, consignmentString);
                var allreclist = Search(database, HPSDK.BaseObjectTypes.Record, allrecString);

                //if (conslist.Count == allreclist.Count)
                if (allreclist.Count == 0)
                {
                    continue;
                }
                else
                {
                    var rec = new HPSDK.Record(database, recuri);
                    var approvalUri = Search(database, HP.HPTRIM.SDK.BaseObjectTypes.ConsignmentApprover, $"record:{recuri}")[0];
                    var approval = new HPSDK.ConsignmentApprover(database, approvalUri);
                    MarkDescriptionByRemove(database, consignmentNo, recuri.ToString(), "Its child aggregation is not ready for disposition.", "", approval.ApprovedOn.ToDateTime());
                    RemoveConsignmentFlag(recuri, consignmentNo, database);

                    var consignment = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));
                    consignment.RemoveRecord(rec);
                    string strHPEUri = GetFieldValue(database, "HPE RM Uri", rec);
                    string reason = GetFieldValue(database, "Classification code", rec) + " child has different consignment.";
                    consignment.UpdateComment = "[Record Removed from Consignment.] " + reason + " | Removed | " + strHPEUri;
                    consignment.Save();
                    var contentList = Search(database, HPSDK.BaseObjectTypes.Record, $"container:{recuri} and consignment:{consignmentNo}");
                    checkRecConsignment(database, contentList, consignmentNo);
                }
            }
        }
        private void UpdateElectronic(HPSDK.Database database, string dataElectronic)
        {
            if (dataElectronic.Length == 0)
                return;

            string[] Electronic = dataElectronic.Split('|');
            UpdateRecDAnumber(database, Electronic);
        }

        private void UpdateRecDAnumber(HPSDK.Database database, string[] Electronics)
        {
            System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
            foreach (var Electronic in Electronics)
            {
                Hashtable A = js.Deserialize<Hashtable>(Electronic);
                string uri = A["Uri"].ToString();
                string DAnumber = A["DAnumber"].ToString();
                try
                {
                    var rec = new HPSDK.Record(database, Convert.ToInt32(uri));
                    rec.SetFieldValue(new HP.HPTRIM.SDK.FieldDefinition(database, "GRS deposit identifier"), new HPSDK.UserFieldValue(DAnumber));
                    rec.Save();
                }
                catch
                {
                }
            }
        }

        private void Archive(HPSDK.Database database, HPSDK.Consignment consignment,string remark)
        {
            consignment.SetArchivistReviewComplete();            
            consignment.Save();
        }

        private string GetEventsDes(HPSDK.History audit)
        {
            string eventDesc = "";
            switch (audit.OtherEventType.ToString().ToLower())
            {
                case "objectadded":
                    eventDesc = audit.ForObjectType + " Added";
                    break;
                case "objectmodified":
                    eventDesc = audit.ForObjectType + " Modified";
                    break;
                case "objectdeleted":
                    eventDesc = audit.ForObjectType + " Deleted";
                    break;
            }
            return eventDesc;
        }

        private void ConsignmentExport(string consignmentNo)
        {
            ExportCore export = new ExportCore();

            try
            {
                export.exportDir = Path.Combine(export.EXPORT_DIR, CONSIGNMENT_FOLDER);
                export.logFileDir = Path.Combine(export.logFileDir, CONSIGNMENT_FOLDER);
                export.SessionID = Session.SessionID;

                string retUri = GetTheRootTreeUriForExport();
                foreach (var uri in retUri.TrimEnd(',').Split(','))
                {
                    string exportFileName = export.ExportAndSaveEletronic($"Uri:{uri}+");
                    StringBuilder sbAudit = new StringBuilder();
                    sbAudit.Append(export.GetSecurityAuditHeader(uri)).AppendLine();
                    sbAudit.Append(export.GetSecurityAuditContent(uri));
                    export.WriteCsv(sbAudit.ToString());
                    string filePath = export.ZipPackage(export.exportDir);
                    export.exportDir = Path.Combine(export.EXPORT_DIR, CONSIGNMENT_FOLDER);
                }
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDB())
                {
                    var consignment = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));
                    string hpeuri = consignment.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).AsString();
                    string actionsRmk = "Exported at "+DateTime.Now;
                    consignment.SetFieldValue(new HPSDK.FieldDefinition(database, "Actions"), new HPSDK.UserFieldValue(actionsRmk));
                    consignment.UpdateComment = "[Consignment Exported] " + consignment.Number + " | Exported | " + hpeuri;
                    consignment.Save();
                }
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(true);
                    JW.WriteMember("message");
                    JW.WriteString("Export successfully.");
                    JW.WriteMember("path");
                    JW.WriteString("The files are exported according to web.config file.");
                    JW.WriteEndObject();
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString("Export failed. " + ex.Message);
                    JW.WriteEndObject();
                }
            }
            finally
            {
                export.CloseDatabase();
            }
        }

        private string GetTheRootTreeUriForExport()
        {
            string consignmentNo = Request["consignmentNo"];

            string searchstring = $"consignment:{consignmentNo}";
            string rootUri = "";
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
            {
                var list = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                foreach (var recUri in list)
                {
                    var rec = new HPSDK.Record(database, recUri);
                    if (rec.RecordType.Name != "Record" && rec.RecordType.Name != "Compound Record")
                    {
                        string ContainerConsignment = rec.Container.ConsignmentObject == null ? "" : rec.Container.ConsignmentObject.Uri.ToString();
                        if (ContainerConsignment == consignmentNo)
                            continue;
                    }

                    if (rec.Contents != "")
                    {
                        rootUri += recUri + ",";
                    }
                }
            }

            return rootUri.TrimEnd(',');
        }

        private void CheckConsignmentAccessByUser()
        {
            string officer = Request["officer"];
            bool hasAccess = false;
            bool status = true;
            string message = "";
            try
            {
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDB())
                {

                    HPSDK.Location loc = new HPSDK.Location(database, officer);
                    hasAccess = loc.HasPermission(HPSDK.UserPermissions.RecordArchivist);
                    //if (loc.TypeOfLocation.ToString() == "Group")
                    //{
                    //    hasAccess = loc.GetFieldValue(new HPSDK.FieldDefinition(database, "CWC-F032-Consignment")).AsBool();
                    //}
                    //else
                    //{
                    //    string searchstring = $"CWCF032Consignment and hasMember:{loc.Uri}";
                    //    var list = Search(database, HPSDK.BaseObjectTypes.Location, searchstring);
                    //    if (list.Count > 0)
                    //        hasAccess = true;

                    //}
                }
            }
            catch (Exception ex)
            {
                status = false;
                message = ex.Message;
            }
            using (JsonWriter JW = CreateJsonWriter(Response.Output))
            {
                JW.WriteStartObject();
                JW.WriteMember("success");
                JW.WriteBoolean(status);
                JW.WriteMember("hasAccess");
                JW.WriteBoolean(hasAccess);
                JW.WriteMember("message");
                JW.WriteString(message);
                JW.WriteEndObject();
            }
        }

        #region report
        private void Report(HPSDK.Database database,string consignmentNo,string method, HPSDK.Consignment consignment, HPSDK.Record review,string reason)
        {
            //List<List<HPSDK.Record>> listAllRecord = GetRecordListByConsignmentNo(database, consignmentNo, method);
            List<Reject> approvalList = GetApproverRejectResonByConsignmentUri(database, consignmentNo, method);
            List<Issues> listIssues = GetIssuesDesc(database, consignmentNo, method);
            PdfPTable tab = null;
            AppendReportTask(ref tab, database, approvalList, method, consignment, review, reason, listIssues);

            //string destDir = System.Configuration.ConfigurationManager.AppSettings["SaffronWorkArea"]+"/TempUpload/"+ Session.SessionID + "/";
            string destDir = Server.MapPath(_stubReportPath);
            var fileName = "CompleteConsignment_InTray_Report_PDF_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            var suffix = ".pdf";
            var filePath = destDir + fileName + suffix;

            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            Document doc = new Document(PageSize.A4.Rotate(), 5, 5, 5, 5);
            PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.Create));

            doc.Open();
            doc.Add(tab);
            doc.Close();

            //var fileContents = System.Text.Encoding.Default.GetBytes(tab.ToString());
            //WriteBuffToFile(fileContents, filePath);

            AddInTrayReport(database, destDir, fileName, suffix);
        }

        private List<Reject> GetApproverRejectResonByConsignmentUri(HPSDK.Database database, string consignmentNo, string method)
        {
            var approvalList = new List<Reject>();
            if (method.ToLower() == "review")
            {
                //select all
                //var searchstring = " type:[default:\"Review consignment\"] and not ConsignmentReviewRecordsCount:0 and not ConsignmentIsComplete and not ConsignmentReviewDate:blank ";
                //HPSDK.TrimURIList uriList = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                HPSDK.Record reviewConsign = new HPSDK.Record(database,new HPSDK.TrimURI(long.Parse(consignmentNo)));
                if (reviewConsign != null)
                {
                    var selStr = $"ConsignmentReviewUri:{consignmentNo}";
                    HPSDK.TrimURIList uriList = Search(database, HPSDK.BaseObjectTypes.Record, selStr);
                    if (uriList != null && uriList.Count > 0)
                    {
                        foreach (var uri in uriList)
                        {
                            var aggregation = new HPSDK.Record(database, uri);
                            if (aggregation != null)
                            {
                                var approDate = aggregation.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Approval Date"));
                                var date = approDate == null ? "" : approDate.ToString();
                                approvalList.Add(new Reject() { recUri = aggregation.Uri.UriAsString, eventDate = date, reason = "" });
                            }
                        }
                    }
                    
                    //status:reject
                    HPSDK.FieldDefinition rejectText = new HPSDK.FieldDefinition(database, "Consignment Reject Text");
                    string rejectJson = reviewConsign.GetFieldValue(rejectText).AsString();
                    if (rejectJson != null && rejectJson != "")
                    {
                        rejectJson = "[" + rejectJson + "]";
                        List<Reject> rejectList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Reject>>(rejectJson);
                        if (rejectList != null && rejectList.Count > 0)
                        {
                            approvalList.AddRange(rejectList);
                        }
                    }
                }
            }
            else
            {
                //var searchstring = $"consignment:{consignmentNo} and status:Approved,\"Rejected(PendingRemoval)\",\"Rejected(Removed)\"";                
                HPSDK.TrimURIList uriList = Search(database, HPSDK.BaseObjectTypes.ConsignmentApprover, $"consignment:{consignmentNo}");
                if (uriList != null && uriList.Count > 0)
                {
                    foreach (var approvaluri in uriList)
                    {
                        var approval = new HPSDK.ConsignmentApprover(database, approvaluri);
                        if (approval != null && approval.Name != "")
                        {
                            if (approval.Status.ToString().ToLower() == "approved")
                            {
                                approvalList.Add(new Reject() { recUri = approval.Record.Uri.UriAsString, eventDate = approval.ApprovedOn.ToString(), reason = "" });
                            }
                        }
                    }
                }
                //rejected
                HPSDK.Consignment consig = new HPSDK.Consignment(database, Convert.ToInt32(consignmentNo));
                if (consig != null)
                {
                    HPSDK.FieldDefinition rejText = new HPSDK.FieldDefinition(database, "Consignment Reject Text");
                    string rejJson = consig.GetFieldValue(rejText).AsString();
                    if (rejJson != "")
                    {
                        rejJson = "[" + rejJson + "]";
                        List<Reject> rejectList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Reject>>(rejJson);
                        if (rejectList != null)
                        {
                            approvalList.AddRange(rejectList);
                        }
                    }
                }
            }            
            return approvalList;
        }

        private List<Issues> GetIssuesDesc(HPSDK.Database database, string consignmentNo, string method)
        {
            var listIssues = new List<Issues>();
            if (method.ToLower() == "review")
            {
                HPSDK.Record reviewConsign = new HPSDK.Record(database,new HPSDK.TrimURI(long.Parse(consignmentNo)));
                if (reviewConsign != null)
                {
                    HPSDK.FieldDefinition issuesText = new HPSDK.FieldDefinition(database, "Consignment Remove Text");
                    var issuesJson = reviewConsign.GetFieldValue(issuesText);
                    if (issuesJson != null && issuesJson.AsString() != "")
                    {
                        var isJson = "[" + issuesJson.ToString()+"]";
                        List<Issues> issuesList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Issues>>(isJson);
                        if (issuesList != null && issuesList.Count > 0)
                        {
                            foreach (var item in issuesList)
                            {
                                item.type = item.description.IndexOf("holds") != -1 ? "holds" : "other" ;
                                listIssues.Add(item);
                            }
                        }
                    }
                }
            }
            else
            {
                HPSDK.Consignment consig = new HPSDK.Consignment(database,Convert.ToInt32(consignmentNo));
                if (consig != null)
                {
                    HPSDK.FieldDefinition txtIssues = new HPSDK.FieldDefinition(database, "Consignment Remove Text");
                    var issuJson = consig.GetFieldValue(txtIssues);
                    if (issuJson != null && issuJson.AsString() != "")
                    {
                        var json = "[" + issuJson.ToString() + "]";
                        List<Issues> issuList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Issues>>(json);
                        if (issuList != null && issuList.Count > 0)
                        {
                            foreach (var item in issuList)
                            {
                                item.type = item.description.IndexOf("holds") != -1 ? "holds" : "other";
                                listIssues.Add(item);
                            }
                        }
                    }
                }
            }
            return listIssues;
        }
        
        private TableCell SetTableCell(string text)
        {
            TableCell cell = new TableCell();
            cell.Text = text;
            return cell;
        }

        public static void WriteBuffToFile(byte[] buff, string filePath)
        {
            string directoryName = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directoryName))
            {
                Directory.CreateDirectory(directoryName);
            }
            FileStream output = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            var encode = System.Text.Encoding.GetEncoding("UTF-8");
            BinaryWriter writer = new BinaryWriter(output, encode);
            writer.Write(buff, 0, buff.Length);
            writer.Flush();
            writer.Close();
            output.Close();
        }

        private void AppendReportTask(ref PdfPTable dt, HPSDK.Database database, List<Reject> approvalList, string method,HPSDK.Consignment consig,HPSDK.Record record,string reason, List<Issues> listIssues)
        {
            try
            {
                #region
                var consignmentName = "";
                var cutoffDate = "";
                var officer = "";
                var createDate = "";
                var sendEmailDate = "";
                var reviewDate = "";
                if (method.ToLower() == "review" && record != null)
                {
                    consignmentName = record.Title;
                    cutoffDate = record.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment CutOffDate")).AsDate().ToDateTime().ToString("yyyy-MM-dd");
                    officer = record.OwnerLocation.SortName;
                    reviewDate = record.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Review Date")).AsDate().ToDateTime().ToString("yyyy-MM-dd HH:mm:ss");
                    var createTime = record.GetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Created")).ToString();
                    createDate = ConvertTDateIime(createTime);
                    string sendTime = record.GetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Sent")).ToString();
                    sendEmailDate = ConvertTDateIime(sendTime);                    
                }
                else {
                    consignmentName = consig.Name;
                    cutoffDate = consig.CutoffDate.ToDateTime().ToString("yyyy-MM-dd");
                    officer = consig.Archivist.SortName;
                    reviewDate = consig.DateReviewed.ToDateTime().ToString("yyyy-MM-dd HH:mm:ss");
                    var createTime = consig.GetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Created")).ToString();
                    createDate = ConvertTDateIime(createTime);
                    string sendTime = consig.GetFieldValue(new HPSDK.FieldDefinition(database, "Date Time Sent")).ToString();
                    sendEmailDate = ConvertTDateIime(sendTime);
                }
                //, "Receive Date"
                string[] str = { "Task", "Performer", "Action", "Complete Date", "Reason" };

                dt = new PdfPTable(str.Length);
                dt.TotalWidth = 800f;
                dt.LockedWidth = true;

                BaseFont arial = BaseFont.CreateFont(@"c:\Windows\Fonts\Arial.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                Font font = new Font(arial, 12);
                Font font15 = new Font(arial,15);

                //add title
                PdfPCell pc = new PdfPCell(new Phrase(consignmentName, font15));
                pc.Colspan = str.Length;
                pc.HorizontalAlignment = 1;
                pc.VerticalAlignment = 1;
                pc.PaddingBottom = 10;
                pc.Border = 0;
                dt.AddCell(pc);

                //add Action Date
                PdfPCell pc1 = new PdfPCell(new Phrase("Action Date: "+ DateTime.Now.ToString("yyyy-MM-dd"), font));
                pc1.Colspan = str.Length;
                pc1.HorizontalAlignment = 0;
                pc1.VerticalAlignment = 1;
                pc1.PaddingBottom = 5;
                pc1.Border = 0;
                dt.AddCell(pc1);
                //add Cutoff Date
                PdfPCell pc2 = new PdfPCell(new Phrase("Cutoff Date: "+cutoffDate, font));
                pc2.Colspan = str.Length;
                pc2.HorizontalAlignment = 0;
                pc2.VerticalAlignment = 1;
                pc2.PaddingBottom = 10;
                pc2.Border = 0;
                dt.AddCell(pc2);

                PdfPCell pc3 = new PdfPCell(new Phrase("Tasks", font15));
                pc3.Colspan = str.Length;
                pc3.HorizontalAlignment = 0;
                pc3.VerticalAlignment = 1;
                pc3.PaddingBottom = 10;
                pc3.Border = 0;
                dt.AddCell(pc3);
                for (int i = 0; i < str.Length; i++)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(str[i], font));
                    cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                    cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    dt.AddCell(cell);
                }
                #endregion
                #region //table
                var listReject = new List<Issues>();
                var listAction = new List<HPSDK.Record>();
                if (TaskList != null && TaskList.Count > 0)
                {
                    foreach (var task in TaskList)
                    {
                        if (task.TaskID == 4)
                        {
                            dt.AddCell(task.TaskName);
                            dt.AddCell("");
                            dt.AddCell(task.TaskAction);
                            //dt.AddCell("");//receive
                            dt.AddCell("");//complete                        
                            dt.AddCell("");

                            #region for part
                            List<Issues> others = listIssues.Where(t => t.type == "other").ToList();
                            if (others != null && others.Count > 0)
                            {
                                foreach (var o in others)
                                {
                                    HPSDK.Record r = new HPSDK.Record(database, new HPSDK.TrimURI(long.Parse(o.recUri)));
                                    if (r != null)
                                    {
                                        dt.AddCell(GetFieldValue(database, "Classification code", r));
                                        dt.AddCell(r.OwnerLocation.Name);
                                        dt.AddCell("Approval");
                                        dt.AddCell(ConvertTDateIime(o.date));
                                        dt.AddCell("");
                                    }
                                }
                            }
                            if (approvalList.Count > 0)
                            {
                                foreach (var recReject in approvalList)
                                {
                                    HPSDK.Record part = new HPSDK.Record(database, new HPSDK.TrimURI(long.Parse(recReject.recUri)));
                                    
                                    if (part != null)
                                    {
                                        dt.AddCell(GetFieldValue(database, "Classification code", part));
                                        dt.AddCell(part.OwnerLocation.Name);
                                        dt.AddCell(recReject.reason == "" ? "Approval" : "Reject");
                                        //dt.AddCell("");//receive date
                                        var rejectDate = recReject.eventDate;
                                        dt.AddCell(ConvertTDateIime(rejectDate));
                                        dt.AddCell(recReject.reason);
                                        if (recReject.reason != "")
                                        {
                                            listReject.Add(new Issues() { recUri = part.Uri.UriAsString, description = recReject.reason, type = "reject" });
                                        }
                                        else {
                                            listAction.Add(part);
                                        }
                                    }
                                }
                            }
                            #endregion
                        }
                        else
                        {
                            dt.AddCell(task.TaskName);
                            dt.AddCell(officer);
                            dt.AddCell(task.TaskAction);
                            //dt.AddCell("");//receive
                            switch (task.TaskID)
                            {
                                case 1:
                                    dt.AddCell(createDate);//complete  
                                    dt.AddCell("");
                                    break;
                                case 2:
                                    dt.AddCell(reviewDate);
                                    dt.AddCell("");
                                    break;
                                case 3:
                                    dt.AddCell(sendEmailDate);
                                    dt.AddCell("");
                                    break;
                                case 5:
                                    dt.AddCell(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                    dt.AddCell(reason);
                                    break;
                            }
                        }
                    }                    
                }
                #endregion

                #region //disposal action 
                if (listIssues.Count > 0)
                {
                    List<Issues> listHold = listIssues.Where(t=>t.type == "holds").ToList();
                    if (listHold != null && listHold.Count > 0)
                    {
                        ShowIssue(database, listHold, "applied disposal hold", dt, str.Length, font15, font);
                    }
                    List<Issues> listOther = listIssues.Where(t => t.type == "other").ToList();
                    if (listOther != null && listOther.Count > 0)
                    {
                        ShowIssue(database,listOther, "others", dt,str.Length,font15,font);
                    }
                }
                if (listReject != null && listReject.Count > 0)
                {
                    ShowIssue(database, listReject, "approver rejected", dt, str.Length, font15, font);
                }
                #endregion
                #region
                if (listAction != null && listAction.Count > 0)
                {
                    string prefix = "Aggregation(s) is/are ready for \" {0} \"";
                    var action = "";
                    foreach (var p in listAction)
                    {
                        string disposalAction = p.RetentionSchedule == null ? "" : p.RetentionSchedule.GetFieldValue(new HPSDK.FieldDefinition(database, "Disposal Action")).AsString();
                        if (action.IndexOf(disposalAction) < 0 && disposalAction != "")
                        {
                            action += disposalAction + ",";
                        }
                    }

                    if (action != "")
                    {
                        var scheduleType = action.TrimEnd(',').Split(',');
                        foreach (var item in scheduleType)
                        {
                            PdfPCell cell = new PdfPCell(new Phrase(string.Format(prefix, item), font15));
                            cell.Colspan = str.Length;
                            cell.HorizontalAlignment = 0;
                            cell.VerticalAlignment = 1;
                            cell.PaddingTop = 10;
                            cell.Border = 0;
                            dt.AddCell(cell);

                            foreach (var rec in listAction)
                            {
                                string disposalAction = rec.RetentionSchedule == null ? "" : rec.RetentionSchedule.GetFieldValue(new HPSDK.FieldDefinition(database, "Disposal Action")).AsString();
                                if (item == disposalAction && disposalAction != "")
                                {
                                    var code = GetFieldValue(database, "Classification path", rec);
                                    PdfPCell cellRecord = new PdfPCell(new Phrase(code, font));
                                    cellRecord.Colspan = str.Length;
                                    cellRecord.HorizontalAlignment = 0;
                                    cellRecord.VerticalAlignment = 1;
                                    cellRecord.PaddingTop = 5;
                                    cellRecord.PaddingLeft = 10;
                                    cellRecord.Border = 0;
                                    dt.AddCell(cellRecord);
                                }
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                GlobalFunc.Log(ex);
            }
        }

        private void ShowIssue(HPSDK.Database database, List<Issues> listIssue,string type,PdfPTable dt,int cellCount,Font font15,Font font)
        {
            string strIssue = string.Format("Aggregation(s) is/are Not ready for Disposition({0})",type);
            PdfPCell cell = new PdfPCell(new Phrase(strIssue, font15));
            cell.Colspan = cellCount;
            cell.HorizontalAlignment = 0;
            cell.VerticalAlignment = 1;
            cell.PaddingTop = 10;
            cell.Border = 0;
            dt.AddCell(cell);

            foreach (var issue in listIssue)
            {
                var issueRec = new HPSDK.Record(database, new HPSDK.TrimURI(long.Parse(issue.recUri)));
                if (issueRec != null)
                {
                    var code = GetFieldValue(database, "Classification path", issueRec);
                    var desc = "(" + issue.description + ")";
                    PdfPCell cellRecord = new PdfPCell(new Phrase(code + "    " + desc, font));
                    cellRecord.Colspan = cellCount;
                    cellRecord.HorizontalAlignment = 0;
                    cellRecord.VerticalAlignment = 1;
                    cellRecord.PaddingTop = 5;
                    cellRecord.PaddingLeft = 10;
                    cellRecord.Border = 0;
                    dt.AddCell(cellRecord);
                }
            }
        }

        private void AddInTrayReport(HPSDK.Database database, string destDir, string fileName, string suffix)
        {
            var record = new HPSDK.Record(database, new HP.HPTRIM.SDK.RecordType(database, "InTray-Report"));
            record.Title = fileName;
            record.Security = "UNCLASSIFIED";

            var locList = SearchLocations(database, GlobalFunc.UserName);
            HPSDK.Location loc = new HPSDK.Location(database, locList[0]);
            record.Assignee = loc;
            
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }
            var filePath = destDir + fileName + suffix;
            //string destPath = System.IO.Path.Combine(destDir);
            TwsLib.Tws tws = new TwsLib.Tws();
            //destPath = tws.PatchAwayExif(destPath, Session.SessionID);
            HPSDK.InputDocument inDoc = new HPSDK.InputDocument();
            //inDoc.SetAsFile(destPath);
            inDoc.SetAsFile(filePath);
            record.SetDocument(inDoc, false, false, "");
            record.Save();

            string strHPEUri = record.Uri.UriAsString;
            record.UpdateComment = "[InTray-Report Added] Genarate Stub Report |Create| " + strHPEUri;
            record.Save();
        }

        private string ConvertTDateIime(string paramTime)
        {
            if (paramTime != "")
            {
                if (paramTime.IndexOf(" at ") != -1) 
                {
                    var date = paramTime.Substring(0, paramTime.IndexOf(" at "));
                    var time = paramTime.Substring(paramTime.IndexOf(" at ") + 3);
                    paramTime = DateTime.Parse(date + time).ToString("yyyy-MM-dd HH:mm:ss");
                }
                else
                {
                    paramTime = DateTime.Parse(paramTime).ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
            return paramTime;
        }
        #endregion

        #region stub
        private void ModifyRecordType(string consignmentNo, string reason)
        {
            using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
            {
                List<List<HPSDK.Record>> listAllRecord = GetRecordListByConsignmentNo(database, consignmentNo);

                if (listAllRecord.Count > 0)
                {
                    foreach (var listPart in listAllRecord)
                    {
                        List<HPSDK.Record> listRecord = new List<HPSDK.Record>();
                        CompleteStub(ref listRecord, database, listPart, consignmentNo, reason, -1, null);
                        //get part => record
                        GetRecordByPartUri(ref listRecord, database, listPart, consignmentNo);

                        if (listRecord.Count > 0)
                        {
                            listRecord.Reverse();
                            foreach (var record in listRecord)
                            {
                                //add log
                                string strHPEUri = record.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                                record.UpdateComment = "[Complete Consignment->Generate Stub] Record Deleted |Deleted| " + strHPEUri;
                                record.Save();

                                record.Refresh();
                                record.Delete();                                
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This function is only for import operation.
        /// </summary>
        /// <param name="listRecord"></param>
        /// <param name="database"></param>
        /// <param name="listPart"></param>
        /// <param name="consignmentNo"></param>
        /// <param name="reason"></param>
        /// <param name="parentUri"></param>
        /// <param name="parentStub"></param>
        /// <returns></returns>
        public string SaveStubForImport(HPSDK.Record rec, RecordParameter recordParameter)
        {
            using(var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
            {
                var record = new HPSDK.Record(new HPSDK.RecordType(database, "Stub"));
                SetFieldValue(record, database, rec, "Consignment Uri Flag", recordParameter.ConsignmentUriFlag);
                //SetFieldValue(record, database, rec, "Uniform resource identifier", rec.Uri, false);//read only
                //SetFieldValue(record, database, rec, "Relation - Entity", "");//
                SetFieldValue(record, database, rec, "Reason", recordParameter.Reason);//

                SetFieldValue(record, database, rec, "Classification code", recordParameter.ClassificationCode);
                SetFieldValue(record, database, rec, "Classification path", recordParameter.ClassificationPath);
                //record.ManualDestructionDate = rec.DisposalDate;//Date disposed 

                //record.OwnerLocation = rec.OwnerLocation;
                SetFieldValue(record, database, rec, "Owner1", recordParameter.SLOC__53);
                SetFieldValue(record, database, rec, "Part Number", recordParameter.PartNumber ?? "11112222");//rec.Number

                SetFieldValue(record, database, rec, "Security Classification Remarks", recordParameter.SecurityClassificationRemarks);
                record.Security = rec.Security;//Security Classification
                SetFieldValue(record, database, rec, "Security classification type", recordParameter.SecurityClassificationType);
                SetFieldValue(record, database, rec, "Stub type", recordParameter.StubType);
                SetFieldValue(record, database, rec, "HPE RM Uri", "");//System identifier
                record.Title = HttpUtility.UrlDecode(recordParameter.Title);
                record.IsEnclosed = true;
                record.Container = rec;
                //var disposalType = rec.ConsignmentObject.DisposalMethod.ToString().ToLower() == "destroy" ? HPSDK.DisposalType.Destroyed : HPSDK.DisposalType.Transferred;
                //record.Dispose(disposalType, false);

                record.Save();

                //if (!string.IsNullOrEmpty(recordParameter.Date_DateClosed))
                //{
                //    record.DateClosed = new HPSDK.TrimDateTime(recordParameter.Date_DateClosed + " " + recordParameter.Time_DateClosed);
                //    record.Save();
                //}

                try
                {
                    record.DateClosed = new HPSDK.TrimDateTime(DateTime.Now);
                    record.Save();
                }
                catch (Exception)
                {
                }

                return record.Number;
            }
        }

        private void CompleteStub(ref List<HPSDK.Record> listRecord, HPSDK.Database database,List<HPSDK.Record> listPart,string consignmentNo,string reason,long parentUri, HPSDK.Record parentStub)
        {
            var pUri = parentUri == -1 ? listPart[0].Container.Uri.Value : parentUri;
            var listTop = listPart.Where(t => t.Container.Uri.Value == pUri).ToList();
            var listOther = listPart.Where(x=>x.Container.Uri.Value != pUri).ToList();

            if (listTop != null && listTop.Count > 0)
            {
                List<HPSDK.Record> listStub = new List<HPSDK.Record>();
                foreach (var rec in listTop)
                {
                    #region
                    var record = new HPSDK.Record(new HPSDK.RecordType(database, "Stub"));
                    SetFieldValue(record, database, rec, "Consignment Uri Flag", consignmentNo);
                    //SetFieldValue(record, database, rec, "Uniform resource identifier", rec.Uri, false);//read only
                    SetFieldValue(record, database, rec, "Relation - Entity", "");
                    SetFieldValue(record, database, rec, "Reason", reason);//

                    SetFieldValue(record, database, rec, "Classification code", "");
                    SetFieldValue(record, database, rec, "Classification path", "");
                    record.ManualDestructionDate = rec.DisposalDate;//Date disposed 
                    
                    record.OwnerLocation = rec.OwnerLocation;
                    SetFieldValue(record, database, rec, "Owner1", rec.OwnerLocation.Name);
                    //record.SetFieldValue(new HPSDK.FieldDefinition(database, "Owner"), new HPSDK.UserFieldValue(rec.OwnerLocation));
                    SetFieldValue(record, database, rec, "Remark", "");
                    SetFieldValue(record, database, rec, "Part Number", "");//rec.Number

                    SetFieldValue(record, database, rec, "Security Classification Remarks", "");
                    record.Security = rec.Security;//Security Classification
                    SetFieldValue(record, database, rec, "Security classification type", "");
                    SetFieldValue(record, database, rec, "Stub type", rec.RecordType.Name);
                    SetFieldValue(record, database, rec, "HPE RM Uri", "");//System identifier
                    record.Title = rec.Title;
                    record.IsEnclosed = true;
                    var disposalType = rec.ConsignmentObject.DisposalMethod.ToString().ToLower() == "destroy" ? HPSDK.DisposalType.Destroyed : HPSDK.DisposalType.Transferred;
                    record.Dispose(disposalType, false);
                    #endregion

                    if (parentUri == -1)
                    {
                        record.Container = rec.Container;
                    }
                    else
                    {
                        record.Container = parentStub;
                    }
                    
                    listStub.Add(record);
                    record.Save();

                    //add log
                    record.DateClosed = new HPSDK.TrimDateTime(DateTime.Now);
                    string strHPEUri = record.GetFieldValue(new HPSDK.FieldDefinition(database, "HPE RM Uri")).ToString();
                    record.UpdateComment = "[Complete Consignment->Generate Stub] Add Stub Type Record |Created| " + strHPEUri;
                    record.Save();

                    //child node is stub,update container
                    var searchstring = $"container:{rec.Uri} and type[\"Stub\"]";
                    var listuri = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                    if (listuri != null && listuri.Count > 0)
                    {
                        foreach (var uri in listuri)
                        {
                            var stubRec = new HPSDK.Record(database, uri);
                            if (stubRec != null)
                            {
                                stubRec.Container = record;
                                stubRec.Save();
                            }
                        }
                    }

                    listRecord.Add(rec);
                }
                for (int i = 0; i < listTop.Count; i++)
                {
                    if (listOther != null && listOther.Count > 0)
                    {
                        CompleteStub(ref listRecord, database, listOther, consignmentNo, reason, listTop[i].Uri, listStub[i]);
                    }
                }
            }            
        }

        private void GetRecordByPartUri(ref List<HPSDK.Record> listRecord,HPSDK.Database database, List<HPSDK.Record> listPart,string consignmentNo)
        {
            if (listPart != null && listPart.Count > 0)
            {
                var parts = listPart.Where(t => t.RecordType.Uri == (int)ErksNodeType.Part).ToList();
                foreach (var rec in parts)
                {
                    var searchstring = $"container:{rec.Uri}";
                    var listuri = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
                    if (listuri != null && listuri.Count > 0)
                    {
                        //select CompoundRecord/Record,Record/Component
                        foreach (var uri in listuri)
                        {
                            var recordByPart = new HPSDK.Record(database, uri);
                            if (recordByPart.RecordType.Uri == (int)ErksNodeType.CompoundRecord || recordByPart.RecordType.Uri == (int)ErksNodeType.Record || recordByPart.RecordType.Uri == (int)ErksNodeType.Component)
                            {
                                if (recordByPart.ConsignmentObject != null)
                                {
                                    var consig = new HPSDK.Consignment(database, new HPSDK.TrimURI(long.Parse(consignmentNo)));
                                    consig.RemoveRecord(recordByPart);
                                }
                                listRecord.Add(recordByPart);
                                if (recordByPart.RecordType.Uri == (int)ErksNodeType.CompoundRecord || recordByPart.RecordType.Uri == (int)ErksNodeType.Record)
                                {
                                    var searchstr = $"container:{recordByPart.Uri}";
                                    var recUri = Search(database, HPSDK.BaseObjectTypes.Record, searchstr);
                                    if (recUri != null && recUri.Count > 0)
                                    {
                                        foreach (var r in recUri)
                                        {
                                            var recByCompoundRec = new HPSDK.Record(database, r);
                                            if (recByCompoundRec != null && recByCompoundRec.RecordType.Uri == (int)ErksNodeType.Record || recByCompoundRec.RecordType.Uri == (int)ErksNodeType.Component)
                                            {
                                                if (recByCompoundRec.ConsignmentObject != null)
                                                {
                                                    var consig = new HPSDK.Consignment(database, new HPSDK.TrimURI(long.Parse(consignmentNo)));
                                                    consig.RemoveRecord(recByCompoundRec);
                                                }
                                                listRecord.Add(recByCompoundRec);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private List<List<HPSDK.Record>> GetRecordListByConsignmentNo(HPSDK.Database database, string consignmentNo)
        {
            List<List<HPSDK.Record>> listAllRecord = new List<List<HPSDK.Record>>();
            //select all record
            string searchstring = $"consignment:{consignmentNo}";
            var list = Search(database, HPSDK.BaseObjectTypes.Record, searchstring);
            if (list.Count > 0)
            {
                List<HPSDK.Record> listRecord = new List<HPSDK.Record>();
                foreach (var item in list)
                {
                    var rec = new HPSDK.Record(database, item);
                    if (rec != null)
                    {
                        var consignmentUriFlag = rec.GetFieldValue(new HPSDK.FieldDefinition(database, "Consignment Uri Flag"));
                        if (consignmentUriFlag != null && consignmentUriFlag.ToString() != "")
                        {
                            listRecord.Add(rec);
                        }
                    }
                }

                //select top record
                List<HPSDK.Record> listTopRecord = new List<HPSDK.Record>();
                string searchstr = $"consignment:{consignmentNo} and not container:[consignment:[{consignmentNo}]]";
                var listTopRec = Search(database, HPSDK.BaseObjectTypes.Record, searchstr);
                if (listTopRec.Count > 0)
                {
                    foreach (var t in listTopRec)
                    {
                        var rec = new HPSDK.Record(database, t);
                        if (rec != null)
                        {
                            listTopRecord.Add(rec);
                        }
                    }
                }

                foreach (var t in listTopRecord)
                {
                    List<HPSDK.Record> listAll = new List<HPSDK.Record>();
                    listAll.Add(t);
                    GetNextRecord(ref listAll, listRecord, t.Uri.UriAsString);
                    listAllRecord.Add(listAll);
                }
            }
            return listAllRecord;
        }

        private void GetNextRecord(ref List<HPSDK.Record> listAll, List<HPSDK.Record> listRec, string topUri)
        {
            if (listRec != null && listRec.Count > 0)
            {
                List<HPSDK.Record> listNext = listRec.Where(t => t.Container.Uri.UriAsString == topUri).ToList();
                if (listNext != null && listNext.Count() > 0)
                {
                    List<HPSDK.Record> listOther = listRec.Where(t => t.Container.Uri.UriAsString != topUri).ToList();
                    foreach (var r in listNext)
                    {
                        listAll.Add(r);
                        if (r.RecordType.Uri != (int)ErksNodeType.Part)
                        {
                            GetNextRecord(ref listAll, listOther, r.Uri.UriAsString);
                        }
                    }
                }
            }
        }

        private void SetFieldValue(HPSDK.Record record, HPSDK.Database database, HPSDK.Record rec, string field, string val)
        {
            if (val == "")
            {
                string fieldVal;
                if (field == "Part Number" && (int)rec.RecordType.Uri != (int)SaffronWebApp.Web.ErksNodeType.Part)
                {
                    fieldVal = "null";
                }
                else
                {
                    fieldVal = GetFieldValue(database, field, rec);
                }
                record.SetFieldValue(new HPSDK.FieldDefinition(database, field), new HPSDK.UserFieldValue(fieldVal));
            }
            else
            {
                record.SetFieldValue(new HPSDK.FieldDefinition(database, field), new HPSDK.UserFieldValue(val));
            }          
        }
        #endregion

        private HPSDK.TrimURIList SearchLocations(HPSDK.Database database, string name)
        {
            string searchstring = $"login:\"{name}\"";
            return Search(database, HPSDK.BaseObjectTypes.Location, searchstring);
        }

        private void GetLoginName()
        {
            try
            {
                using (var database = SaffronWebApp.CodeLib.TwsNewSDK.connectDBAsAdmin())
                {
                    HPSDK.TrimURIList name = SearchLocations(database, GlobalFunc.UserName);
                    if (name != null && name.Count >0  )
                    {
                        HPSDK.Location loc = new HPSDK.Location(database, name[0]);
                        if (loc != null)
                        {
                            using (JsonWriter JW = CreateJsonWriter(Response.Output))
                            {
                                JW.WriteStartObject();
                                JW.WriteMember("success");
                                JW.WriteBoolean(true);
                                JW.WriteMember("UserName");
                                JW.WriteString(loc.SortName);
                                JW.WriteEndObject();
                            }
                        }                        
                    }                    
                }
            }
            catch (Exception ex)
            {
                using (JsonWriter JW = CreateJsonWriter(Response.Output))
                {
                    JW.WriteStartObject();
                    JW.WriteMember("success");
                    JW.WriteBoolean(false);
                    JW.WriteMember("message");
                    JW.WriteString(ex.Message.ToString());
                    JW.WriteEndObject();
                }
            }
        }

        public static string GetIcon(ref HPSDK.Record rec, string Extension)
        {
            string icon = "";

            switch (Extension.ToLower())
            {
                case "pdf":
                    icon = "pdf";
                    break;
                case "doc":
                    icon = "doc";
                    break;
                case "dot":
                    icon = "doc";
                    break;
                case "dotx":
                    icon = "doc";
                    break;
                case "docx":
                    icon = "doc";
                    break;
                case "txt":
                    icon = "txt";
                    break;
                case "rtf":
                    icon = "txt";
                    break;
                case "xls":
                    icon = "xls";
                    break;
                case "xlsx":
                    icon = "xls";
                    break;
                case "ppt":
                case "pptx":
                case "pps":
                    icon = "ppt";
                    break;
                case "gif":
                case "jpeg":
                case "jpg":
                    icon = "image";
                    break;
                case "tif":
                case "tiff":
                    icon = "tif";
                    break;
                case "bmp":
                case "png":
                    icon = "image";
                    break;
                case "zip":
                case "rar":
                    icon = "zip";
                    break;
                case "rmvb":
                case "rm":
                    icon = "real";
                    break;
                case "mepg":
                case "3gp":
                case "asf":
                case "avi":
                case "flv":
                case "dat":
                case "wmv":
                    icon = "video";
                    break;
                case "swf":
                    icon = "flash";
                    break;
                case "eml":
                    icon = "email";
                    break;
                case "htm":
                case "html":
                case "mht":
                    icon = "htm";
                    break;
                case "mp3":
                case "mp4":
                case "wma":
                    icon = "mp3";
                    break;
                case "vmbx":
                    try
                    {
                        if (rec.HasEmailAttachments)
                        {
                            icon = "emailAttachment";
                        }
                        else
                        {
                            icon = "email";
                        }
                    }
                    catch (Exception ex)
                    {
                        GlobalFunc.Log(ex);
                        icon = "attachment";
                    }
                    break;

                default:
                    icon = "attachment";
                    break;
            }
            return icon;
        }

        private static string GetFolderIcon(HPSDK.Record rec, HPSDK.Database database)
        {
            string FolderTypeIcon = "";
            if ((int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.Folder || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.SubFolder || (int)rec.RecordType.Uri == (int)SaffronWebApp.Web.ErksNodeType.Part)
            {
                HPSDK.FieldDefinition FolderType = new HPSDK.FieldDefinition(database, "Folder Type");
                Object FoldeTypeVal = rec.GetFieldValue(FolderType);
                if (FoldeTypeVal != null)
                {
                    switch (FoldeTypeVal.ToString())
                    {
                        case "Electronic":
                            FolderTypeIcon = "electronic";
                            break;
                        case "Hybrid":
                            FolderTypeIcon = "hybrid";
                            break;
                        case "Physical":
                            FolderTypeIcon = "physical";
                            break;

                    }
                }
            }
            return FolderTypeIcon;
        }

        public static string GetRecordIcon(HPSDK.RecordType RecType)
        {
            int id = (Int32)RecType.TrimIconId;

            if (id >= 502 && id <= 574)
            {
                return "rticon_" + id;
            }
            else
            {
                return "rticon_511";
            }
        }

        private static JsonWriter CreateJsonWriter(TextWriter writer)
        {
            JsonTextWriter jsonWriter = new JsonTextWriter(writer);
            jsonWriter.PrettyPrint = true;
            return jsonWriter;
        }

        private void ExportStub(long consignmentUri)
        {
        }
        
        public List<Task> TaskList {
            get {
                List<Task> tasklist = new List<Task>();
                tasklist.Add(new Task() { TaskID = 1, TaskName = "Create Consignment", TaskAction = "Create" });
                tasklist.Add(new Task() { TaskID = 2, TaskName = "Complete Review", TaskAction = "Review" });
                tasklist.Add(new Task() { TaskID = 3, TaskName = "Send Email to Approver(s)", TaskAction = "Send Email" });
                tasklist.Add(new Task() { TaskID = 4, TaskName = "Seeking Approval", TaskAction = "Show Approval" });
                tasklist.Add(new Task() { TaskID = 5, TaskName = "Perform Disposal Action", TaskAction = "Complete" });
                return tasklist;
            }
        }
    }
    public class Reject {
        public string recUri { get; set; }
        public string eventDate { get; set; }
        public string reason { get; set; }
    }
    public class Issues
    {
        public string recUri { get; set; }
        public string date { get;set; }
        public string description { get; set; }
        public string type { get; set; }
    }
    public class Task {
        public int TaskID { get; set;}
        public string TaskName { get; set; }
        public string TaskAction { get; set; }
    }
}