﻿using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HonorariumController : ControllerBase
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

        public HonorariumController(IConfiguration configuration)
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


        [HttpPost("AddHonorariumData")]
        public IActionResult AddHonorariumData(HonorariumPaymentList formData)
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;

                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

                StringBuilder addedHcpData = new StringBuilder();
                int addedHcpDataNo = 1;
                CultureInfo hindi = new CultureInfo("hi-IN");
                foreach (var i in formData.HcpRoles)
                {
                    string rowData = $"{addedHcpDataNo}. Name:{i.HcpName} | Role:{i.HcpRole} |Code:{i.MisCode} | HCP Type:{i.GOorNGO}| Including GST:{i.IsInclidingGst}| Agreement Amount:{i.AgreementAmount} ";

                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                }
                string panalist = addedHcpData.ToString();
                var newRow = new Row();
                newRow.Cells = new List<Cell>();

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = formData.RequestHonorariumList.slideKits });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.RequestHonorariumList.EventId });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Type"), Value = formData.RequestHonorariumList.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Date"), Value = formData.RequestHonorariumList.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Topic"), Value = formData.RequestHonorariumList.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.RequestHonorariumList.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.RequestHonorariumList.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Start Time"), Value = formData.RequestHonorariumList.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "End Time"), Value = formData.RequestHonorariumList.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Venue Name"), Value = formData.RequestHonorariumList.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel & Accommodation Amount"), Value = formData.RequestHonorariumList.TotalTravelAndAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Honorarium Amount"), Value = formData.RequestHonorariumList.TotalHonorariumSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Budget"), Value = formData.RequestHonorariumList.TotalSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = formData.RequestHonorariumList.Expenses });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel Amount"), Value = formData.RequestHonorariumList.TotalTravelSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Accommodation Amount"), Value = formData.RequestHonorariumList.TotalAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Expense"), Value = formData.RequestHonorariumList.TotalExpenses });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Local Conveyance"), Value = formData.RequestHonorariumList.TotalLocalConveyance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = formData.RequestHonorariumList.Brands });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = formData.RequestHonorariumList.Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = formData.RequestHonorariumList.Panelists });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.RequestHonorariumList.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists & Agreements"), Value = panalist });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Honorarium Submitted?"), Value = formData.RequestHonorariumList.HonarariumSubmitted });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.RequestHonorariumList.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM"), Value = formData.RequestHonorariumList.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.RequestHonorariumList.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator"), Value = formData.RequestHonorariumList.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Compliance"), Value = formData.RequestHonorariumList.Compliance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts"), Value = formData.RequestHonorariumList.FinanceAccounts });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.FinanceTreasury });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager"), Value = formData.RequestHonorariumList.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "1 Up Manager"), Value = formData.RequestHonorariumList.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.RequestHonorariumList.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formData.RequestHonorariumList.Role });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRow });
                var eventId = formData.RequestHonorariumList.EventId;
                var x = 1;
                foreach (var p in formData.RequestHonorariumList.Files)
                {
                    var name = " AttachedFile";
                    var filePath = SheetHelper.testingFile(p, eventId, name);
                    var addedRow = addedRows[0];
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }



                if (formData.RequestHonorariumList.IsDeviationUpload == "Yes")
                {
                    try
                    {
                        var newRow7 = new Row();
                        newRow7.Cells = new List<Cell>();
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.RequestHonorariumList.EventTopic });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.RequestHonorariumList.EventType });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.RequestHonorariumList.EventDate });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.RequestHonorariumList.StartTime });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.RequestHonorariumList.EndTime });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.RequestHonorariumList.VenueName });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.RequestHonorariumList.City });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.RequestHonorariumList.State });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium documentation not uploaded within 5 working days" });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HON-5Workingdays Deviation Date Trigger"), Value = formData.RequestHonorariumList.IsDeviationUpload });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.RequestHonorariumList.Compliance });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.RequestHonorariumList.Compliance });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.RequestHonorariumList.InitiatorName });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.RequestHonorariumList.InitiatorEmail });

                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });


                        var j = 1;
                        foreach (var p in formData.RequestHonorariumList.DeviationFiles)
                        {
                            var name = "2WorkingDaysAfterEvent";
                            var filePath = SheetHelper.testingFile(p, eventId, name);
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
                    catch (Exception ex)
                    {
                        return BadRequest(ex.Message);
                    }
                }

                var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
                if (targetRow != null)
                {
                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Honorarium Submitted?");
                    var cellToUpdateB = new Cell
                    { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                    var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }
                    smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                }
                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }



    }
}
