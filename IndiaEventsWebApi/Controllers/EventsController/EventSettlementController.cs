﻿using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
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
        private readonly SmartsheetClient smartsheet;


        public EventSettlementController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        }

        [HttpPost("AddEventSettlementData")]
        public IActionResult AddEventSettlementData(EventSettlement formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

                StringBuilder addedInviteesData = new StringBuilder();
                StringBuilder addedExpences = new StringBuilder();
                int addedInviteesDataNo = 1;
                int addedExpencesNo = 1;
                foreach (var formdata in formData.expenseSheets)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
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

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.EventId });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventTopic"), Value = formData.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventType"), Value = formData.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventDate"), Value = formData.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "StartTime"), Value = formData.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EndTime"), Value = formData.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "VenueName"), Value = formData.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Attended"), Value = formData.Attended });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InviteesParticipated"), Value = Invitee });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "ExpenseDetails"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "TotalExpenseDetails"), Value = formData.TotalExpense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "AdvanceDetails"), Value = formData.Advance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InitiatorName"), Value = formData.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = formData.Brands });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = Invitee });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = formData.Panalists });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = formData.SlideKits });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Invitees"), Value = formData.totalInvitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Attendees"), Value = formData.TotalAttendees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IsAdvanceRequired"), Value = formData.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PostEventSubmitted?"), Value = formData.PostEventSubmitted });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Head"), Value = formData.FinanceHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance"), Value = formData.FinanceHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM"), Value = formData.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator"), Value = formData.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.MarkeringHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Compliance"), Value = formData.Compliance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts"), Value = formData.FinanceAccounts });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.FinanceTreasury });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager"), Value = formData.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "1 Up Manager"), Value = formData.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formData.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel Amount"), Value = formData.TotalTravelSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Accommodation Amount"), Value = formData.TotalAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Expense"), Value = formData.TotalExpenses });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel & Accommodation Amount"), Value = formData.TotalTravelAndAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Honorarium Amount"), Value = formData.TotalHonorariumSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Budget"), Value = formData.TotalSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Actual"), Value = formData.TotalActuals });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Advance Utilized For Event"), Value = formData.AdvanceUtilizedForEvents });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Pay Back Amount To Company"), Value = formData.PayBackAmountToCompany });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Additional Amount Needed To Pay For Initiator"), Value = formData.AdditionalAmountNeededToPayForInitiator });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Local Conveyance"), Value = formData.TotalLocalConveyance });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                var x = 1;
                foreach (var p in formData.Files)
                {

                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];

                    var name = r.Split(".")[0];

                    var filePath = SheetHelper.testingFile(q, val, name);




                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (formData.EventOpen30Days == "Yes" || formData.EventLessThan5Days == "Yes" || formData.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formData.DeviationFiles)
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

                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = val });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.EndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.State });
                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST- Beyond45Days Deviation Date Trigger"), Value = formData.EventOpen30Days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                            }
                            else if (file == "Lessthan5InviteesDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Lessthan5Invitees Deviation Trigger"), Value = formData.EventLessThan5Days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Less than 5 attendees excluding speaker" });
                            }
                            else if (file == "ExcludingGSTDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Excluding GST?"), Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "POST-Deviation Excluding GST" });
                            }



                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.Compliance });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.Compliance });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.InitiatorEmail });

                            var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });




                            var j = 1;
                            foreach (var p in formData.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                var r = words[0];
                                var q = words[1];
                                if (deviationname == r)
                                {
                                    var name = r.Split(".")[0];

                                    var filePath = SheetHelper.testingFile(q, val, name);




                                    var addedRow = addeddeviationrow[0];
                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
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
                return Ok(new { Message = "Data added successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}