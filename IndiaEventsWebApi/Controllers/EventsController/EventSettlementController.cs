﻿using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Diagnostics;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;


namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EventSettlementController : ControllerBase
    {


        private readonly string accessToken;
        private readonly IConfiguration configuration;
        // private readonly SmartsheetClient smartsheet;

        private readonly string panelSheet;
        private readonly string InviteeSheet;
        private readonly string ExpenseSheet;
        private readonly string SlideKitSheet;
        private readonly SemaphoreSlim _externalApiSemaphore;

        public EventSettlementController(IConfiguration configuration, SemaphoreSlim externalApiSemaphore)
        {
            this.configuration = configuration;
            this._externalApiSemaphore = externalApiSemaphore;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

            panelSheet = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            InviteeSheet = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            ExpenseSheet = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            SlideKitSheet = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;

            //smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            //sheet1 = SheetHelper.GetSheetById(smartsheet, panelSheet);
            //sheet2 = SheetHelper.GetSheetById(smartsheet, InviteeSheet);
            //sheet3 = SheetHelper.GetSheetById(smartsheet, ExpenseSheet);
            //sheet4 = SheetHelper.GetSheetById(smartsheet, SlideKitSheet);
        }

        //[HttpPost("AddEventSettlementData")]
        //public IActionResult AddEventSettlementData(EventSettlement formData)
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        //        string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
        //        string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
        //        string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
        //        Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
        //        Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //        Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

        //        StringBuilder addedInviteesData = new StringBuilder();
        //        StringBuilder addedExpences = new StringBuilder();
        //        int addedInviteesDataNo = 1;
        //        int addedExpencesNo = 1;
        //        foreach (var formdata in formData.expenseSheets)
        //        {
        //            string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
        //            addedExpences.AppendLine(rowData);
        //            addedExpencesNo++;
        //        }
        //        string Expense = addedExpences.ToString();
        //        foreach (var formdata in formData.Invitee)
        //        {
        //            string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName} | {formdata.MISCode} | {formdata.LocalConveyance}";
        //            addedInviteesData.AppendLine(rowData);
        //            addedInviteesDataNo++;
        //        }
        //        string Invitee = addedInviteesData.ToString();

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.EventId });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventTopic"), Value = formData.EventTopic });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventType"), Value = formData.EventType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventDate"), Value = formData.EventDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "StartTime"), Value = formData.StartTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EndTime"), Value = formData.EndTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "VenueName"), Value = formData.VenueName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Attended"), Value = formData.Attended });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InviteesParticipated"), Value = Invitee });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "ExpenseDetails"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "TotalExpenseDetails"), Value = int.Parse(formData.TotalExpense) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "AdvanceDetails"), Value = formData.Advance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InitiatorName"), Value = formData.InitiatorName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = formData.Brands });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = Invitee });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = formData.Panalists });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = formData.SlideKits });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Invitees"), Value = int.Parse(formData.totalInvitees )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Attendees"), Value = int.Parse(formData.TotalAttendees )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IsAdvanceRequired"), Value = formData.IsAdvanceRequired });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PostEventSubmitted?"), Value = formData.PostEventSubmitted });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Head"), Value = formData.FinanceHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance"), Value = formData.FinanceHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM"), Value = formData.RBMorBM });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator"), Value = formData.SalesCoordinatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.MarkeringHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Compliance"), Value = formData.Compliance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts"), Value = formData.FinanceAccounts });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.FinanceTreasury });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager"), Value = formData.ReportingManagerEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "1 Up Manager"), Value = formData.FirstLevelEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formData.Role });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel Amount"), Value = int.Parse(formData.TotalTravelSpend) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Accommodation Amount"), Value = int.Parse(formData.TotalAccomodationSpend) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Expense"), Value = int.Parse(formData.TotalExpenses )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel & Accommodation Amount"), Value = int.Parse(formData.TotalTravelAndAccomodationSpend )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Honorarium Amount"), Value = int.Parse(formData.TotalHonorariumSpend) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Budget"), Value = int.Parse(formData.TotalSpend )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Actual"), Value = int.Parse(formData.TotalActuals )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Advance Utilized For Event"), Value = int.Parse(formData.AdvanceUtilizedForEvents) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Pay Back Amount To Company"), Value = int.Parse(formData.PayBackAmountToCompany) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Additional Amount Needed To Pay For Initiator"), Value = int.Parse(formData.AdditionalAmountNeededToPayForInitiator )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Local Conveyance"), Value = int.Parse(formData.TotalLocalConveyance) });


        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRow });

        //        var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId");
        //        var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
        //        var val = eventIdCell.DisplayValue;

        //        var x = 1;
        //        foreach (var p in formData.Files)
        //        {

        //            string[] words = p.Split(':');
        //            var r = words[0];
        //            var q = words[1];

        //            var name = r.Split(".")[0];

        //            var filePath = SheetHelper.testingFile(q, val, name);




        //            var addedRow = addedRows[0];

        //            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                    sheet.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //            x++;


        //            if (System.IO.File.Exists(filePath))
        //            {
        //                SheetHelper.DeleteFile(filePath);
        //            }
        //        }
        //        if (formData.EventOpen30Days == "Yes" || formData.EventLessThan5Days == "Yes" || formData.IsDeviationUpload == "Yes")
        //        {
        //            List<string> DeviationNames = new List<string>();
        //            foreach (var p in formData.DeviationFiles)
        //            {
        //                string[] words = p.Split(':');
        //                var r = words[0];
        //                DeviationNames.Add(r);
        //            }
        //            foreach (var deviationname in DeviationNames)
        //            {
        //                var file = deviationname.Split(".")[0];
        //                var eventId = val;

        //                try
        //                {

        //                    var newRow7 = new Row();
        //                    newRow7.Cells = new List<Cell>();

        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = val });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.EventTopic });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.EventType });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.EventDate });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.StartTime });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.EndTime });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.VenueName });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.City });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.State });
        //                    if (file == "30DaysDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST- Beyond45Days Deviation Date Trigger"), Value = formData.EventOpen30Days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
        //                    }
        //                    else if (file == "Lessthan5InviteesDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Lessthan5Invitees Deviation Trigger"), Value = formData.EventLessThan5Days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Less than 5 attendees excluding speaker" });
        //                    }
        //                    else if (file == "ExcludingGSTDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Excluding GST?"), Value = "Yes" });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "POST-Deviation Excluding GST" });
        //                    }



        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.Compliance });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.FinanceHead });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.InitiatorName });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.InitiatorEmail });

        //                    var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });




        //                    var j = 1;
        //                    foreach (var p in formData.DeviationFiles)
        //                    {
        //                        string[] words = p.Split(':');
        //                        var r = words[0];
        //                        var q = words[1];
        //                        if (deviationname == r)
        //                        {
        //                            var name = r.Split(".")[0];

        //                            var filePath = SheetHelper.testingFile(q, val, name);




        //                            var addedRow = addeddeviationrow[0];
        //                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                            j++;
        //                            if (System.IO.File.Exists(filePath))
        //                            {
        //                                SheetHelper.DeleteFile(filePath);
        //                            }
        //                        }
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    return BadRequest(ex.Message);
        //                }
        //            }
        //        }
        //        return Ok(new { Message = "Data added successfully." });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}



        [HttpPost("AddEventSettlementDataForFirst5"), DisableRequestSizeLimit]
        public async Task<IActionResult> AddEventSettlementDataForFirst5(EventSettlement formData)
        {
            string strMessage = string.Empty;

            strMessage += "==starting of api== " + DateTime.Now + "==";
            try
            {
                Log.Information("Start of EventSettlement Post" + DateTime.Now);
                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));

                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
                string UI_URL = configuration.GetSection("SmartsheetSettings:UI_URL").Value;

                Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);
                strMessage += "==UI_URL Get Completed==" + DateTime.Now.ToString() + "==";
                //long.TryParse(sheetId, out long parsedSheetId);
                //long.TryParse(sheetId1, out long parsedSheetId1);
                //long.TryParse(sheetId7, out long parsedSheetId7);

                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                strMessage += "==EventSettlement Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                strMessage += "==EventRequestProcess Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
                strMessage += "==Deviation_Process Get Completed==" + DateTime.Now.ToString() + "==";
                StringBuilder addedInviteesData = new();
                StringBuilder addedExpences = new();
                int addedInviteesDataNo = 1;
                int addedExpencesNo = 1;
                foreach (var formdata in formData.expenseSheets)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | Budget Amount: {formdata.BudgetAmount}| Actual Amount: {formdata.ActualAmount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                }
                string Expense = addedExpences.ToString();
                foreach (var formdata in formData.Invitee)
                {
                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName} | {formdata.MISCode} | {formdata.LocalConveyance}";

                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }
                string Invitee = addedInviteesData.ToString();

                Dictionary<string, long> Sheetcolumns = new();
                foreach (var column in sheet.Columns)
                {
                    Sheetcolumns.Add(column.Title, (long)column.Id);
                }

                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };
                Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Accounts URL"));
                Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                Row? targetRow3 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Post Event URL"));
                Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));

                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Accounts URL"], Value = targetRow1?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Treasury URL"], Value = targetRow2?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Approver Post Event URL"], Value = targetRow3?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Initiator URL"], Value = targetRow4?.Cells[1].Value ?? "no url" });

                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Pay Back Amount To Company"], Value = SheetHelper.NumCheck(formData.PayBackAmountToCompany) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Additional Amount Needed To Pay For Initiator"], Value = SheetHelper.NumCheck(formData.AdditionalAmountNeededToPayForInitiator) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Local Conveyance"], Value = SheetHelper.NumCheck(formData.TotalLocalConveyance) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["EventId/EventRequestId"], Value = formData.EventId });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Event Topic"], Value = formData.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Event Type"], Value = formData.EventType });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Event Date"], Value = formData.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Start Time"], Value = formData.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["End Time"], Value = formData.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Venue Name"], Value = formData.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["City"], Value = formData.City });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["State"], Value = formData.State });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Attended"], Value = SheetHelper.NumCheck(formData.Attended) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Invitees Participated"], Value = Invitee });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Expense Details"], Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Expense Details"], Value = formData.TotalExpense });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Advance Details"], Value = formData.Advance });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Initiator Name"], Value = formData.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Brands"], Value = formData.Brands });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Invitees"], Value = Invitee });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Advance Amount"], Value = formData.AdvanceProvided });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Expense BTC"], Value = formData.TotalBtcAmount });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Expense BTE"], Value = formData.TotalBteAmount });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Panelists"], Value = formData.Panalists });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["SlideKits"], Value = formData.SlideKits });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Expenses"], Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Invitees"], Value = SheetHelper.NumCheck(formData.totalInvitees) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Attendees"], Value = SheetHelper.NumCheck(formData.TotalAttendees) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["IsAdvanceRequired"], Value = formData.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["PostEventSubmitted?"], Value = formData.PostEventSubmitted });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Medical Affairs Head"], Value = formData.MedicalAffairsHead });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Head"], Value = formData.FinanceHead });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance"], Value = formData.FinanceHead });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Initiator Email"], Value = formData.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["RBM/BM"], Value = formData.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Sales Head"], Value = formData.SalesHead });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Sales Coordinator"], Value = formData.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Marketing Coordinator"], Value = formData.MarketingCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Marketing Head"], Value = formData.MarkeringHead });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Compliance"], Value = formData.Compliance });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Accounts"], Value = formData.FinanceAccounts });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Treasury"], Value = formData.FinanceTreasury });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Reporting Manager"], Value = formData.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["1 Up Manager"], Value = formData.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Role"], Value = formData.Role });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Travel Amount"], Value = SheetHelper.NumCheck(formData.TotalTravelSpend) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Accommodation Amount"], Value = SheetHelper.NumCheck(formData.TotalAccomodationSpend) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Expense"], Value = SheetHelper.NumCheck(formData.TotalExpenses) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Travel & Accommodation Amount"], Value = SheetHelper.NumCheck(formData.TotalTravelAndAccomodationSpend) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Honorarium Amount"], Value = SheetHelper.NumCheck(formData.TotalHonorariumSpend) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Budget Amount"], Value = formData.TotalBudgetAmount });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Actual"], Value = SheetHelper.NumCheck(formData.TotalActuals) });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Advance Utilized For Event"], Value = SheetHelper.NumCheck(formData.AdvanceUtilizedForEvents) });

                // IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRow });

                strMessage += "== Before EventSettlement  Add in sheet==" + DateTime.Now.ToString() + "==";
                IList<Row> addedRows = ApiCalls.DeviationData(smartsheet, sheet, newRow);
                strMessage += "==After EventSettlement  Add in sheet==" + DateTime.Now.ToString() + "==";

                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;
                int x = 1;
                foreach (var p in formData.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, name);
                    Row addedRow = addedRows[0];
                    strMessage += "==Before Add Attachment in sheet==" + x.ToString() + DateTime.Now.ToString() + "==";
                    Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet, addedRow, filePath);
                    strMessage += "==After Add Attachment in sheet==" + x.ToString() + DateTime.Now.ToString() + "==";
                    // Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (formData.EventOpen30Days == "Yes" || formData.EventLessThan5Days == "Yes" || formData.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    //foreach (var p in formData.DeviationFiles)
                    //{
                    //    string[] words = p.Split(':');
                    //    string r = words[0];
                    //    DeviationNames.Add(r);
                    //}
                    foreach (var p in formData.DeviationFiles)
                    {

                        string[] words = p.Split(':')[0].Split("*");
                        string r = words[1];
                        DeviationNames.Add(r);
                    }
                    Dictionary<string, long> Sheet7columns = new();
                    foreach (var column in sheet7.Columns)
                    {
                        Sheet7columns.Add(column.Title, (long)column.Id);
                    }
                    foreach (var deviationname in DeviationNames)
                    {
                        string file = deviationname.Split(".")[0];
                        string eventId = val;
                        try
                        {
                            Row newRow7 = new()
                            {
                                Cells = new List<Cell>()
                            };
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventId/EventRequestId"], Value = val });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Topic"], Value = formData.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Type"], Value = formData.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Date"], Value = formData.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Start Time"], Value = formData.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["End Time"], Value = formData.EndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Venue Name"], Value = formData.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["City"], Value = formData.City });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["State"], Value = formData.State });
                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST- Beyond45Days Deviation Date Trigger"], Value = formData.EventOpen30Days });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInEventSettlement:30DaysDeviationFile").Value });
                            }
                            else if (file == "Lessthan5InviteesDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Lessthan5Invitees Deviation Trigger"], Value = formData.EventLessThan5Days });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInEventSettlement:Lessthan5InviteesDeviationFile").Value });
                            }
                            else if (file == "ExcludingGSTDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Deviation Excluding GST?"], Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInEventSettlement:ExcludingGSTDeviationFile").Value });
                            }
                            else if (file == "Change in venue")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Deviation Change in venue trigger"], Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = file });

                            }
                            else if (file == "Change in topic")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Deviation Change in topic trigger"], Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = file });

                            }
                            else if (file == "Change in speaker")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Deviation Change in speaker trigger"], Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = file });

                            }
                            else if (file == "Attendees not captured in photographs")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Deviation Attendees not captured trigger"], Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = file });

                            }
                            else if (file == "Speaker not captured in photographs")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Deviation Speaker not captured trigger"], Value = "Yes" });//POST-Deviation Speaker not captured  trigger
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = file });
                            }
                            else
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["POST-Deviation Other Deviation Trigger"], Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = "Others" });
                                newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Other Deviation Type"], Value = file });
                            }
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Head"], Value = formData.SalesHead });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Finance Head"], Value = formData.FinanceHead });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Name"], Value = formData.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Email"], Value = formData.InitiatorEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Coordinator"], Value = formData.SalesCoordinatorEmail });

                            //IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });
                            strMessage += "==Before Add Data in deviation sheet==" + x.ToString() + DateTime.Now.ToString() + "==";
                            IList<Row> addeddeviationrow = ApiCalls.DeviationData(smartsheet, sheet7, newRow7);
                            strMessage += "==After Add Data in deviation sheet==" + x.ToString() + DateTime.Now.ToString() + "==";
                            int j = 1;
                            foreach (var p in formData.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                string r = words[0].Split("*")[1];
                                string q = words[1];
                                if (deviationname == r)
                                {
                                    string name = words[0].Split("*")[0];
                                    string filePath = SheetHelper.testingFile(q, name);
                                    Row addedRow = addeddeviationrow[0];
                                    strMessage += "==Before Add Attachment in deviation sheet=="+ DateTime.Now.ToString() + "==";
                                    Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet7, addedRow, filePath);
                                    strMessage += "==After Add Attachment in deviation sheet==" + x.ToString() + DateTime.Now.ToString() + "==";
                                    strMessage += "==Before Add Attachment in Eventsettlement sheet==" + x.ToString() + DateTime.Now.ToString() + "==";
                                    Attachment attachmentintoMain = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet, addedRows[0], filePath);
                                    strMessage += "==After Add Attachment in Eventsettlement sheet==" + DateTime.Now.ToString() + "==";


                                    //Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    //Attachment attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");
                                    j++;
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }
                            }
                            //}
                            //catch (Exception ex)
                            //{
                            //    //return BadRequest(ex.Message);
                            //    Log.Error($"Error occured on EventSettlementController method {ex.Message} at {DateTime.Now}");
                            //    Log.Error(ex.StackTrace);
                            //    return BadRequest(ex.Message);
                            //}
                        }
                        catch (Exception ex)
                        {

                            return BadRequest(new
                            {
                                Message = ex.Message + "------" + ex.StackTrace
                            });
                        }
                    }
                }
                Row s = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role");
                Cell UpdateB = new Cell { ColumnId = ColumnId, Value = formData.Role };
                Row updateRows = new Row { Id = s.Id, Cells = new Cell[] { UpdateB } };
                Cell? cellsToUpdate = s.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formData.Role; }

                strMessage += "==Before Update Role in Eventsettlement sheet==" + DateTime.Now.ToString() + "==";
                smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRows });
                strMessage += "==After Update Role in Eventsettlement sheet==" + DateTime.Now.ToString() + "==";
                Log.Information("End of EventSettlement Post" + DateTime.Now);
                strMessage += "==End Of Api==" + DateTime.Now.ToString() + "==";
               //return Ok(strMessage);
                return Ok(new
                { Message = "Data added successfully." });


            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on UpdateClassIPreEvent method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });
            }
        }

        [HttpPut("EventSettlementUpdate")]
        public async Task<IActionResult> EventSettlementUpdate(UpdateAttendees formData)
        {
            string strMessage = string.Empty;

            strMessage += "==starting of api== " + DateTime.Now + "==";
            try
            {
                Log.Information("Start of EventSettlement Update" + DateTime.Now);

                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));
               
                if (formData.InviteesData.Count > 0)
                {

                    Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, InviteeSheet);
                    strMessage += "==Invitee sheet Get Completed==" + DateTime.Now.ToString() + "==";
                    foreach (var formdata in formData.InviteesData)
                    {
                        var targetRow = sheet2.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.InviteeId));
                        if (targetRow != null)
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };

                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Attended?"), Value = formdata.IsAttended });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Actual Local Conveyance Amount"), Value = formdata.ActualAmount });
                            strMessage += "==Before updtate data in invitee==" + DateTime.Now.ToString() + "==";
                            IList<Row> updatedRow = await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet2, updateRow));
                            strMessage += "==After updtate data in invitee==" + DateTime.Now.ToString() + "==";
                            // IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet2.Id.Value, new Row[] { updateRow });
                            if (formdata.IsUploadDocument == "Yes")
                            {
                                strMessage += "==Before GetAttachmentsFromSheet invitee==" + DateTime.Now.ToString() + "==";
                                PaginatedResult<Attachment> attachments = await Task.Run(() => ApiCalls.GetAttachmantsFromSheet(smartsheet, sheet2, targetRow));
                                strMessage += "==After GetAttachmentsFromSheet invitee==" + DateTime.Now.ToString() + "==";

                                //PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet2.Id.Value, targetRow.Id.Value, null);

                                if (attachments.Data != null || attachments.Data.Count > 0)
                                {

                                    foreach (var attachment in attachments.Data)
                                    {
                                        long Id = attachment.Id.Value;
                                        strMessage += "==Before DeleteAttachment==" + DateTime.Now.ToString() + "==";
                                        smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                                          sheet2.Id.Value,           // sheetId
                                          Id            // attachmentId
                                        );
                                        strMessage += "==After DeleteAttachment==" + DateTime.Now.ToString() + "==";

                                    }

                                }
                                foreach (var p in formdata.UploadDocument)
                                {
                                    string[] words = p.Split(':');
                                    var r = words[0];
                                    var q = words[1];
                                    var name = r.Split(".")[0];
                                    var filePath = SheetHelper.testingFile(q, name);
                                    var addedRow = updatedRow[0];
                                    //var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    //        sheet2.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    strMessage += "==Before AddAttachmentsToSheet==" + DateTime.Now.ToString() + "==";
                                    Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet2, addedRow, filePath);
                                    strMessage += "==After AddAttachmentsToSheet==" + DateTime.Now.ToString() + "==";
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

                            }
                        }
                    }
                }
                if (formData.PanelData.Count > 0)
                {
                    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, panelSheet);
                    strMessage += "==After updtate data in panel sheet==" + DateTime.Now.ToString() + "==";
                    foreach (var formdata in formData.PanelData)
                    {
                        var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.PanelId));
                        if (targetRow != null)
                        {

                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Attended?"), Value = formdata.IsAttended });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Accomodation Amount"), Value = formdata.ActualAccomodationAmount });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Travel Amount"), Value = formdata.ActualTravelAmount });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Local Conveyance Amount"), Value = formdata.ActualLCAmount });
                            strMessage += "==Before updtate panel==" + DateTime.Now.ToString() + "==";
                            IList<Row> updatedRow = await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRow));
                            strMessage += "==After updtate panel==" + DateTime.Now.ToString() + "==";

                            //IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                            if (formdata.IsUploadDocument == "Yes")
                            {
                                PaginatedResult<Attachment> attachments = await Task.Run(() => ApiCalls.GetAttachmantsFromSheet(smartsheet, sheet1, targetRow));


                                // PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, targetRow.Id.Value, null);
                                if (attachments.Data != null || attachments.Data.Count > 0)
                                {

                                    foreach (var attachment in attachments.Data)
                                    {
                                        var name = attachment.Name;
                                        if (name.ToLower().Contains("invoice"))
                                        {
                                            long Id = attachment.Id.Value;
                                            strMessage += "==Before DeleteAttachment==" + DateTime.Now.ToString() + "==";
                                            smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                                              sheet1.Id.Value,           // sheetId
                                              Id            // attachmentId
                                            );
                                            strMessage += "==After DeleteAttachment==" + DateTime.Now.ToString() + "==";
                                        }
                                    }

                                }
                                foreach (var p in formdata.UploadDocument)
                                {
                                    string[] words = p.Split(':');
                                    var r = words[0];
                                    var q = words[1];
                                    var name = r.Split(".")[0];
                                    var filePath = SheetHelper.testingFile(q, name);
                                    var addedRow = updatedRow[0];
                                    //var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    //        sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    strMessage += "==Before AddAttachmentsToSheet==" + DateTime.Now.ToString() + "==";
                                    Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRow, filePath);
                                    strMessage += "==After AddAttachmentsToSheet==" + DateTime.Now.ToString() + "==";
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

                            }
                        }
                    }
                }
                if (formData.ExpenseData.Count > 0)
                {
                    foreach (var formdata in formData.ExpenseData)
                    {
                        Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, ExpenseSheet);
                        strMessage += "==Loded ExpenseSheet==" + DateTime.Now.ToString() + "==";
                        var targetRow = sheet3.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.ExpenseId));
                        if (targetRow != null)
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Actual Amount"), Value = formdata.ActualAmount });
                            strMessage += "==Before updtate data in ExpenseSheet==" + DateTime.Now.ToString() + "==";
                            IList<Row> updatedRow = await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet3, updateRow));
                            strMessage += "==After updtate data in ExpenseSheet==" + DateTime.Now.ToString() + "==";
                            // IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet3.Id.Value, new Row[] { updateRow });
                            if (formdata.IsUploadDocument == "Yes")
                            {
                                PaginatedResult<Attachment> attachments = await Task.Run(() => ApiCalls.GetAttachmantsFromSheet(smartsheet, sheet3, targetRow));


                                //PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet3.Id.Value, targetRow.Id.Value, null);
                                if (attachments.Data != null || attachments.Data.Count > 0)
                                {

                                    foreach (var attachment in attachments.Data)
                                    {
                                        long Id = attachment.Id.Value;
                                        strMessage += "==Before DeleteAttachment==" + DateTime.Now.ToString() + "==";
                                        smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                                          sheet3.Id.Value,           // sheetId
                                          Id            // attachmentId
                                        );
                                        strMessage += "==After DeleteAttachment==" + DateTime.Now.ToString() + "==";

                                    }

                                }
                                foreach (var p in formdata.UploadDocument)
                                {

                                    string[] words = p.Split(':');
                                    var r = words[0];
                                    var q = words[1];
                                    var name = r.Split(".")[0];
                                    var filePath = SheetHelper.testingFile(q, name);
                                    var addedRow = updatedRow[0];
                                    //var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    //        sheet3.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    strMessage += "==Before AddAttachmentsToSheet==" + DateTime.Now.ToString() + "==";
                                    Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet3, addedRow, filePath);
                                    strMessage += "==After AddAttachmentsToSheet==" + DateTime.Now.ToString() + "==";
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

                            }
                        }
                    }
                }
                //if (formData.SlideKitData.Count > 0)
                //{
                //    foreach (var formdata in formData.SlideKitData)
                //    {
                //        Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, SlideKitSheet);
                //        var targetRow = sheet4.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.SlideKitId));
                //        if (targetRow != null)
                //        {
                //            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                //            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.ProductName });
                //            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.IndicationsDone });
                //            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.BatchNumber });
                //            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.SubjectNameandSurName });

                //            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet4.Id.Value, new Row[] { updateRow });
                //            if (formdata.IsUploadDocument == "Yes")
                //            {
                //                foreach (var p in formdata.UploadDocument)
                //                {

                //                    string[] words = p.Split(':');
                //                    var r = words[0];
                //                    var q = words[1];
                //                    var name = r.Split(".")[0];
                //                    var filePath = SheetHelper.testingFile(q, name);
                //                    var addedRow = updatedRow[0];
                //                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                //                            sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                //                    if (System.IO.File.Exists(filePath))
                //                    {
                //                        SheetHelper.DeleteFile(filePath);
                //                    }
                //                }

                //            }
                //        }
                //    }
                //}
                Log.Information("End of EventSettlement Update" + DateTime.Now);
                strMessage += "==End of Api==" + DateTime.Now.ToString() + "==";
                //return Ok(strMessage);
                return Ok(new { Message = "Attendees Updated Successfully" });
                //}
                //catch (Exception ex)
                //{
                //    Log.Error($"Error occured on EventSettlementController method {ex.Message} at {DateTime.Now}");
                //    Log.Error(ex.StackTrace);
                //    return BadRequest(ex.Message);
                //}
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on UpdateClassIPreEvent method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });
            }

        }


    }
}
