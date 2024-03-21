using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers.EventsController
{
    [Route("api/[controller]")]
    [ApiController]
    public class AllPreEventsController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        //private readonly Sheet sheet1;
        //private readonly Sheet sheet2;
        //private readonly Sheet sheet3;
        //private readonly Sheet sheet4;
        //private readonly Sheet sheet5;
        //private readonly Sheet sheet6;
        //private readonly Sheet sheet7;

        public AllPreEventsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            //string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            //string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            //string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            //string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            //string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            //string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            //string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;           
            //sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            //sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            //sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            //sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            //sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            //sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            //sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
        }

        [HttpPost("ClassIPreEvent"), DisableRequestSizeLimit]
        public IActionResult ClassIPreEvent(AllObjModels formDataList)
        {

            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            var TotalHonorariumAmount = 0;
            var TotalTravelAmount = 0;
            var TotalAccomodateAmount = 0;
            var TotalHCPLcAmount = 0;
            var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;

            CultureInfo hindi = new CultureInfo("hi-IN");

            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = int.Parse(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
            }
            string Expense = addedExpences.ToString();
            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {

                string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();

            foreach (var formdata in formDataList.RequestBrandsList)
            {

                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.EventRequestInvitees)
            {
                // string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code: {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";
                string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                addedInviteesData.AppendLine(rowData);
                addedInviteesDataNo++;
                TotalInviteesLcAmount = TotalInviteesLcAmount + int.Parse(formdata.LcAmount);
            }
            string Invitees = addedInviteesData.ToString();

            foreach (var formdata in formDataList.EventRequestHcpRole)
            {
                var HM = int.Parse(formdata.HonarariumAmount);
                var x = string.Format(hindi, "{0:#,#}", HM);
                var t = int.Parse(formdata.Travel) + int.Parse(formdata.Accomdation);
                var y = string.Format(hindi, "{0:#,#}", t);
                //string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {x} |Trav.&Acc.Amt: {y} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + int.Parse(formdata.HonarariumAmount);
                TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.Travel);
                TotalAccomodateAmount = TotalAccomodateAmount + int.Parse(formdata.Accomdation);
                TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();
            var FormattedTotalHonorariumAmount = string.Format(hindi, "{0:#,#}", TotalHonorariumAmount);
            var FormattedTotalTravelAmount = string.Format(hindi, "{0:#,#}", TotalTravelAmount);
            var FormattedTotalAccomodateAmount = string.Format(hindi, "{0:#,#}", TotalAccomodateAmount);
            var FormattedTotalHCPLcAmount = string.Format(hindi, "{0:#,#}", TotalHCPLcAmount);
            var FornattedTotalInviteesLcAmount = string.Format(hindi, "{0:#,#}", TotalInviteesLcAmount);
            var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);
            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var FormattedTotalLC = string.Format(hindi, "{0:#,#}", c);
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

            var FormattedTotal = string.Format(hindi, "{0:#,#}", total);
            var s = TotalTravelAmount + TotalAccomodateAmount;
            var FormattedTotalTAAmount = string.Format(hindi, "{0:#,#}", s);


            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.class1.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.class1.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.class1.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.class1.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.class1.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.class1.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.class1.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.class1.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.class1.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.class1.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.class1.EventWithin7days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.class1.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.class1.AdvanceAmount )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.class1.TotalExpenseBTC )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.class1.TotalExpenseBTE )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.class1.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.class1.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.class1.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.class1.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.class1.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.class1.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.class1.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.class1.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.class1.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.class1.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.class1.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.class1.Role });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                var x = 1;
                foreach (var p in formDataList.class1.Files)
                {
                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];
                    var name = r.Split(".")[0];
                    var filePath = SheetHelper.testingFile(q, val, name);
                    var addedRow = addedRows[0];
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                           sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;

                    // ////////////////////////////////////////////////
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }

                //var targetRow = addedRows[0];
                //long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                //var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.class1.Role };
                //Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                //var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                //if (cellToUpdate != null) { cellToUpdate.Value = formDataList.class1.Role; }

                //smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                if (formDataList.class1.EventOpen30days == "Yes" || formDataList.class1.EventWithin7days == "Yes" || formDataList.class1.FB_Expense_Excluding_Tax == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.class1.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        var r = words[0];

                        DeviationNames.Add(r);
                    }
                    foreach (var deviationname in DeviationNames)
                    {
                        var file = deviationname.Split(".")[0];
                        var eventId = val;
                        try
                        {
                            var newRow7 = new Row();
                            newRow7.Cells = new List<Cell>();

                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.class1.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.class1.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.class1.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.class1.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.class1.EndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.class1.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.class1.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.class1.State });


                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.class1.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.class1.EventOpen30dayscount });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.class1.EventWithin7days });

                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.class1.Sales_Head });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.class1.FinanceHead });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.class1.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.class1.Initiator_Email });

                            var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                            var j = 1;
                            foreach (var p in formDataList.class1.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                var r = words[0];
                                var q = words[1];
                                if (deviationname == r)
                                {
                                    var name = r.Split(".")[0];
                                    var filePath = SheetHelper.testingFile(q, val, name);
                                    var addedRow = addeddeviationrow[0];
                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    j++;
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            return BadRequest(ex.Message);
                        }
                    }
                }


                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.Travel });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.Accomdation });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyance });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.HonarariumAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.class1.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.class1.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.class1.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.class1.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.class1.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.Currency });
                    if (formData.Currency == "Others")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.OtherCurrencyType });
                    }
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BeneficiaryName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.PanNumber });
                    if (formData.HcpRole == "Others")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.OthersType });
                    }
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.PresentationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.PanelSessionPreperationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PanelDisscussionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QASessionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.BriefingSession });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalSessionHours });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });

                    smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                }

                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });
                    smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });
                }

                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    var newRow3 = new Row();
                    newRow3.Cells = new List<Cell>();
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LcAmount });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                    if (formdata.InviteedFrom == "Others")
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType });
                    }
                    else
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode });
                    }
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.class1.EventTopic });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.class1.EventType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.class1.VenueName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.class1.EventDate });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.class1.EventDate });

                    smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, new Row[] { newRow3 });
                }

                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    var newRow5 = new Row();
                    newRow5.Cells = new List<Cell>();
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });

                    smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                }

                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.class1.EventTopic });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.class1.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.class1.VenueName });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.class1.EventDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.class1.EventDate });

                    smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
                }



                var targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.class1.Role };
                Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.class1.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                return Ok(new
                { Message = " Success!" });

            }

            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


        [HttpPost("WebinarPreEvent"), DisableRequestSizeLimit]
        public IActionResult WebinarPreEvent(WebinarPayload formDataList)
        {


            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();
            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            var TotalHonorariumAmount = 0;
            var TotalTravelAmount = 0;
            var TotalAccomodateAmount = 0;
            var TotalHCPLcAmount = 0;
            var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;
            CultureInfo hindi = new CultureInfo("hi-IN");
            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = int.Parse(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {
                string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();

            foreach (var formdata in formDataList.RequestBrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.EventRequestInvitees)
            {
                string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                addedInviteesData.AppendLine(rowData);
                addedInviteesDataNo++;
                TotalInviteesLcAmount = TotalInviteesLcAmount + int.Parse(formdata.LcAmount);
            }
            string Invitees = addedInviteesData.ToString();


            foreach (var formdata in formDataList.EventRequestHcpRole)
            {
                var HM = int.Parse(formdata.HonarariumAmount);
                var x = string.Format(hindi, "{0:#,#}", HM);
                var t = int.Parse(formdata.Travel) + int.Parse(formdata.Accomdation);
                var y = string.Format(hindi, "{0:#,#}", t);
                //string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {x} |Trav.&Acc.Amt: {y} ";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + int.Parse(formdata.HonarariumAmount);
                TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.Travel);
                TotalAccomodateAmount = TotalAccomodateAmount + int.Parse(formdata.Accomdation);
                TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();
            var FormattedTotalHonorariumAmount = string.Format(hindi, "{0:#,#}", TotalHonorariumAmount);
            var FormattedTotalTravelAmount = string.Format(hindi, "{0:#,#}", TotalTravelAmount);
            var FormattedTotalAccomodateAmount = string.Format(hindi, "{0:#,#}", TotalAccomodateAmount);
            var FormattedTotalHCPLcAmount = string.Format(hindi, "{0:#,#}", TotalHCPLcAmount);
            var FornattedTotalInviteesLcAmount = string.Format(hindi, "{0:#,#}", TotalInviteesLcAmount);
            var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);
            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var FormattedTotalLC = string.Format(hindi, "{0:#,#}", c);
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;
            var FormattedTotal = string.Format(hindi, "{0:#,#}", total);
            var s = TotalTravelAmount + TotalAccomodateAmount;
            var FormattedTotalTAAmount = string.Format(hindi, "{0:#,#}", s);
            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.Webinar.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.Webinar.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.Webinar.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.Webinar.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Meeting Type"), Value = formDataList.Webinar.Meeting_Type });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.Webinar.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.Webinar.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.Webinar.EventWithin7days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.Webinar.AdvanceAmount )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.Webinar.TotalExpenseBTC )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.Webinar.TotalExpenseBTE )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.Webinar.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.Webinar.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.Webinar.Marketing_Head });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.Webinar.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.Webinar.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.Webinar.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.Webinar.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.Webinar.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.Webinar.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.Webinar.Role });
                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;
                var x = 1;
                foreach (var p in formDataList.Webinar.Files)
                {
                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];
                    var name = r.Split(".")[0];
                    var filePath = SheetHelper.testingFile(q, val, name);
                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }


                if (formDataList.Webinar.EventOpen30days == "Yes" || formDataList.Webinar.EventWithin7days == "Yes" || formDataList.Webinar.FB_Expense_Excluding_Tax == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.Webinar.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        var r = words[0];

                        DeviationNames.Add(r);
                    }
                    foreach (var deviationname in DeviationNames)
                    {
                        var file = deviationname.Split(".")[0];
                        var eventId = val;
                        try
                        {
                            var newRow7 = new Row();
                            newRow7.Cells = new List<Cell>();
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.Webinar.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.Webinar.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.Webinar.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.Webinar.EndTime });
                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.Webinar.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.Webinar.EventOpen30dayscount });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.Webinar.EventWithin7days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.Webinar.FB_Expense_Excluding_Tax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                            }
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.Webinar.FinanceHead });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });

                            var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });
                            var j = 1;
                            foreach (var p in formDataList.Webinar.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                var r = words[0];
                                var q = words[1];
                                if (deviationname == r)
                                {
                                    var name = r.Split(".")[0];

                                    var filePath = SheetHelper.testingFile(q, val, name);

                                    var addedRow = addeddeviationrow[0];

                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    j++;
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            return BadRequest(ex.Message);
                        }
                    }

                }



                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.Travel });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.Accomdation });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyance });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.HonarariumAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.Webinar.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.Webinar.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.Webinar.EventEndDate });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });

                    if (formData.HcpRole == "Others")
                    {

                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.OthersType });
                    }

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.PresentationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.PanelSessionPreperationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PanelDisscussionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QASessionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.BriefingSession });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalSessionHours });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.Currency });
                    if (formData.Currency == "Others")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.OtherCurrencyType });
                    }
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BeneficiaryName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.PanNumber });


                    smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });







                }







                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });


                    smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

                }
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    var newRow3 = new Row();
                    newRow3.Cells = new List<Cell>();
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LcAmount });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.Webinar.EventType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.Webinar.EventDate });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.Webinar.EventDate });
                    // newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Start Time"), Value = formDataList.Webinar.StartTime });

                    if (formdata.InviteedFrom == "Others")
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType });
                    }
                    else
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode });
                    }
                    smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, new Row[] { newRow3 });
                }
                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    var newRow5 = new Row();
                    newRow5.Cells = new List<Cell>();

                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });


                    smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                }
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.Webinar.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.Webinar.EventDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.Webinar.EventDate });

                    smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
                }


                var targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.Webinar.Role };
                Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.Webinar.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });




                return Ok(new
                { Message = " Success!" });
            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


        [HttpPost("StallFabricationPreEvent"), DisableRequestSizeLimit]
        public IActionResult StallFabricationPreEvent(AllStallFabrication formDataList)
        {
            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            // string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            // sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            // sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            // Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            // Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            // Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            var TotalExpenseAmount = 0;
            CultureInfo hindi = new CultureInfo("hi-IN");

            var uploadDeviationForTableContainsData = !string.IsNullOrEmpty(formDataList.StallFabrication.TableContainsDataUpload) ? "Yes" : "No";
            var EventWithin7Days = !string.IsNullOrEmpty(formDataList.StallFabrication.EventWithin7daysUpload) ? "Yes" : "No";
            var BrouchereUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.EventBrouchereUpload) ? "Yes" : "No";
            var InvoiceUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.Invoice_QuotationUpload) ? "Yes" : "No";


            foreach (var formdata in formDataList.ExpenseSheets)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = int.Parse(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

            }
            string Expense = addedExpences.ToString();


            foreach (var formdata in formDataList.EventBrands)
            {

                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);

            var total = TotalExpenseAmount;

            var FormattedTotal = string.Format(hindi, "{0:#,#}", total);


            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.StallFabrication.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Class III Event Code"), Value = formDataList.StallFabrication.Class_III_EventCode });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.StallFabrication.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.StallFabrication.AdvanceAmount )});

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.StallFabrication.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.StallFabrication.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.StallFabrication.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.StallFabrication.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.StallFabrication.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.StallFabrication.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.StallFabrication.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.StallFabrication.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.StallFabrication.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.StallFabrication.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.StallFabrication.TotalExpenseBTC )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.StallFabrication.TotalExpenseBTE )});

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;



                if (BrouchereUpload == "Yes")
                {
                    var name = " Brochure";
                    var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (InvoiceUpload == "Yes")
                {
                    var name = " Invoice_Quotation";

                    var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }


                if (formDataList.StallFabrication.IsDeviationUpload == "Yes")
                {
                    var eventId = val;
                    List<string> list = new List<string>
                    {
                        "30days","7days"
                    };
                    foreach (var item in list)
                    {
                        if (uploadDeviationForTableContainsData == "Yes" && item == "30days")
                        {
                            try
                            {
                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = uploadDeviationForTableContainsData });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.StallFabrication.EventOpen30dayscount });

                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                var name = "30DaysDeviationFile";
                                var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

                                var addedRow = addeddeviationrow[0];

                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");

                                if (System.IO.File.Exists(filePath))
                                {
                                    SheetHelper.DeleteFile(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                        else if (EventWithin7Days == "Yes" && item == "7days")
                        {
                            try
                            {

                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                                //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.StallFabrication.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                                //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen30days"), Value = uploadDeviationForTableContainsData });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });


                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                var name = "7DaysDeviationFile";
                                var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);



                                var addedRow = addeddeviationrow[0];
                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                                if (System.IO.File.Exists(filePath))
                                {
                                    SheetHelper.DeleteFile(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                    }

                }


                foreach (var formdata in formDataList.EventBrands)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });

                    smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

                }


                foreach (var formdata in formDataList.ExpenseSheets)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
                    //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.StallFabrication.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.StallFabrication.StartDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.StallFabrication.EndDate });

                    smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
                }


                var targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.StallFabrication.Role };
                Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.StallFabrication.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });



                return Ok(new
                { Message = " Success!" });



            }



            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }












        }


        [HttpPost("HCPConsultantPreEvent"), DisableRequestSizeLimit]
        public IActionResult HCPConsultantPreEvent(HCPConsultantPayload formDataList)
        {

            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            //string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            //string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);


            //Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            //Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            //Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            //Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
            //Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            var TotalHonorariumAmount = 0;
            var TotalTravelAmount = 0;
            var TotalAccomodateAmount = 0;
            var TotalHCPLcAmount = 0;
            var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;
            var TotalRegstAmount = 0;


            CultureInfo hindi = new CultureInfo("hi-IN");


            var EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
            var EventWithin7Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventWithin7days) ? "Yes" : "No";
            var BrouchereUpload = !string.IsNullOrEmpty(formDataList.HcpConsultant.BrochureFile) ? "Yes" : "No";
            var FCPA = !string.IsNullOrEmpty(formDataList.HcpConsultant.FcpaFile) ? "Yes" : "No";


            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | RegstAmount: {formdata.RegstAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;

                var amount = int.Parse(formdata.ExpenseAmount);
                var regst = int.Parse(formdata.RegstAmount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
                TotalRegstAmount = TotalRegstAmount + regst;
            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.BrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.HcpList)
            {
                var HM = int.Parse(formdata.RegistrationAmount);
                var x = string.Format(hindi, "{0:#,#}", HM);
                var t = int.Parse(formdata.TravelAmount) + int.Parse(formdata.AccomAmount);
                var y = string.Format(hindi, "{0:#,#}", t);
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} | Regst.Amt: {x} |Trav.&Acc.Amt: {y} ";



                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + int.Parse(formdata.RegistrationAmount);
                TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.TravelAmount);
                TotalAccomodateAmount = TotalAccomodateAmount + int.Parse(formdata.AccomAmount);
                TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LcAmount);
            }
            string HCP = addedHcpData.ToString();
            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount;
            //var BTE = int.Parse(formDataList.HcpConsultant.TotalExpenseBTE);
            //var BTC = int.Parse(formDataList.HcpConsultant.TotalExpenseBTC);

            //var total = BTC + BTE;

            var s = TotalTravelAmount + TotalAccomodateAmount;

            try
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.HcpConsultant.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sponsorship Society Name"), Value = formDataList.HcpConsultant.SponsorshipSocietyName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Country"), Value = formDataList.HcpConsultant.Country });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.HcpConsultant.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.HcpConsultant.AdvanceAmount )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalHonorariumAmount });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.HcpConsultant.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.HcpConsultant.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.HcpConsultant.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.HcpConsultant.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.HcpConsultant.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.HcpConsultant.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.HcpConsultant.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.HcpConsultant.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.HcpConsultant.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.HcpConsultant.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.HcpConsultant.TotalExpenseBTC )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.HcpConsultant.TotalExpenseBTE )});

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                if (BrouchereUpload == "Yes")
                {


                    var filename = "Brochure";
                    var filePath = SheetHelper.testingFile(formDataList.HcpConsultant.BrochureFile, val, filename);


                    var addedRow = addedRows[0];
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }



                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || formDataList.HcpConsultant.AggregateDeviation == "Yes")
                {
                    List<string> DeviationNames = new List<string>();

                    DeviationNames.Add($"30DaysDeviationFile:{formDataList.HcpConsultant.EventOpen30days}");
                    DeviationNames.Add($"7DaysDeviationFile:{formDataList.HcpConsultant.EventWithin7days}");
                    DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.HcpConsultant.AggregateDeviationFiles}");
                    var eventId = val;
                    foreach (var name in DeviationNames)
                    {
                        var y = name.Split(':');
                        var fn = y[0];
                        var bs = y[1];

                        if (bs != "")
                        {
                            try
                            {

                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.HcpConsultant.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.HcpConsultant.EventOpen30dayscount });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 5,00,000 Trigger"), Value = formDataList.HcpConsultant.AggregateDeviation });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 5,00,000 is Exceeded" });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.HcpConsultant.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });

                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                if (fn == "AgregateSpendDeviationFile")
                                {
                                    foreach (var file in formDataList.HcpConsultant.AggregateDeviationFiles)
                                    {
                                        var filename = fn;
                                        var filePath = SheetHelper.testingFile(file, val, filename);
                                        var addedRow = addeddeviationrow[0];
                                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            SheetHelper.DeleteFile(filePath);
                                        }
                                    }
                                }
                                else
                                {
                                    var filename = fn;
                                    var filePath = SheetHelper.testingFile(bs, val, filename);

                                    var addedRow = addeddeviationrow[0];

                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                    }


                }







                foreach (var formData in formDataList.HcpList)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.TravelAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.HcpConsultant.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.HcpConsultant.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.AccomAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LcAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Registration Amount"), Value = formData.RegistrationAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.BudgetAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });



                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });







                    var addeddatarow = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });

                    var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;
                    if (FCPA == "Yes")
                    {
                        var filename = " FCPA";
                        var filePath = SheetHelper.testingFile(formDataList.HcpConsultant.FcpaFile, value, filename);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    }
                }

                foreach (var formdata in formDataList.BrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });


                    smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

                }


                foreach (var formdata in formDataList.ExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Registration Amount"), Value = formdata.RegstAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.ExpenseAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.HcpConsultant.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.HcpConsultant.VenueName });

                    smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
                }




                var targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.HcpConsultant.Role };
                Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.HcpConsultant.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                return Ok(new
                { Message = " Success!" });
            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


        [HttpPost("MedicalUtilityPreEvent"), DisableRequestSizeLimit]
        public IActionResult MedicalUtilityPreEvent(MedicalUtilityPreEventPayload formDataList)
        {


            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            //string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            //string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);



            //string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            //string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            //string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            //string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            //string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;



            //long.TryParse(sheetId1, out long parsedSheetId1);
            //long.TryParse(sheetId2, out long parsedSheetId2);
            //long.TryParse(sheetId4, out long parsedSheetId4);
            //long.TryParse(sheetId6, out long parsedSheetId6);
            //long.TryParse(sheetId7, out long parsedSheetId7);

            //Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            //Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            //Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            //Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
            //Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedHcpDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;


            var TotalExpenseAmount = 0;

            CultureInfo hindi = new CultureInfo("hi-IN");

            //var EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
            var EventOpen30Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventOpen30daysFile) ? "Yes" : "No";
            var EventWithin7Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventWithin7daysFile) ? "Yes" : "No";
            var UploadDeviationFile = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.UploadDeviationFile) ? "Yes" : "No";
            var FCPA = "";


            //if (formDataList.MedicalUtilityData.EventWithin7daysFile != "")
            //{
            //    EventWithin7Days = "Yes";
            //}
            //else
            //{
            //    EventWithin7Days = "No";
            //}
            //if (formDataList.MedicalUtilityData.EventOpen30daysFile != "")
            //{
            //    EventOpen30Days = "Yes";
            //}
            //else
            //{
            //    EventOpen30Days = "No";
            //}
            //if (formDataList.MedicalUtilityData.UploadDeviationFile != "")
            //{
            //    UploadDeviationFile = "Yes";
            //}
            //else
            //{
            //    UploadDeviationFile = "No";
            //}




            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | TotalAmount: {formdata.TotalExpenseAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = int.Parse(formdata.TotalExpenseAmount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

            }
            string Expense = addedExpences.ToString();



            foreach (var formdata in formDataList.BrandsList)
            {

                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();




            foreach (var formdata in formDataList.HcpList)
            {
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} |Speciality: {formdata.Speciality} |Tier: {formdata.Tier} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;

            }
            string HCP = addedHcpData.ToString();



            var total = TotalExpenseAmount;






            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid From"), Value = formDataList.MedicalUtilityData.ValidFrom });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid To"), Value = formDataList.MedicalUtilityData.ValidTill });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.MedicalUtilityData.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.MedicalUtilityData.AdvanceAmount )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.MedicalUtilityData.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.MedicalUtilityData.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.MedicalUtilityData.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.MedicalUtilityData.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.MedicalUtilityData.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.MedicalUtilityData.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.MedicalUtilityData.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.MedicalUtilityData.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.MedicalUtilityData.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.MedicalUtilityData.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.MedicalUtilityData.TotalExpenseBTC )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.MedicalUtilityData.TotalExpenseBTE )});


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;


                List<string> DeviationNames = new List<string>();
                DeviationNames.Add($"30DaysDeviationFile:{formDataList.MedicalUtilityData.EventOpen30daysFile}");
                DeviationNames.Add($"7DaysDeviationFile:{formDataList.MedicalUtilityData.EventWithin7daysFile}");
                DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.MedicalUtilityData.UploadDeviationFile}");

                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || UploadDeviationFile == "Yes")
                {
                    var eventId = val;
                    foreach (var name in DeviationNames)
                    {
                        var y = name.Split(':');
                        var fn = y[0];
                        var bs = y[1];

                        if (bs != "")
                        {
                            try
                            {
                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.MedicalUtilityData.EventOpen30dayscount });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });

                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 1,00,000 Trigger"), Value = UploadDeviationFile });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 1,00,000 is Exceeded" });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.MedicalUtilityData.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });


                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                var filename = fn;
                                var filePath = SheetHelper.testingFile(bs, val, filename);



                                var addedRow = addeddeviationrow[0];

                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");

                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                    }


                }





                foreach (var formData in formDataList.HcpList)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Cost"), Value = int.Parse(formData.MedicalUtilityCostAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Request Date"), Value = formData.HCPRequestDate });

                    //Request Date,fcpa date,Expense type
                    var addeddatarows = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });


                    var FCPAFile = !string.IsNullOrEmpty(formData.UploadFCPA) ? "Yes" : "No";
                    var UploadWrittenRequestDate = !string.IsNullOrEmpty(formData.UploadWrittenRequestDate) ? "Yes" : "No";
                    var Invoice_Brouchere_Quotation = !string.IsNullOrEmpty(formData.Invoice_Brouchere_Quotation) ? "Yes" : "No";



                    var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;

                    if (FCPAFile == "Yes")
                    {

                        var filename = " FCPA";
                        var filePath = SheetHelper.testingFile(formData.UploadFCPA, val, filename);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                    if (UploadWrittenRequestDate == "Yes")
                    {
                        var filename = " UploadWrittenRequestDate";
                        var filePath = SheetHelper.testingFile(formData.UploadWrittenRequestDate, val, filename);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }

                    if (Invoice_Brouchere_Quotation == "Yes")
                    {
                        var filename = " Invoice_Brouchere_Quotation";
                        var filePath = SheetHelper.testingFile(formData.Invoice_Brouchere_Quotation, val, filename);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }





                }

                foreach (var formdata in formDataList.BrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });

                    smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

                }


                foreach (var formdata in formDataList.ExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "MisCode"), Value = formdata.MisCode });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.TotalExpenseAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill });

                    smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
                }



                var targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.MedicalUtilityData.Role };
                Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.MedicalUtilityData.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                return Ok(new
                { Message = " Success!" });
            }

            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


    }
}
