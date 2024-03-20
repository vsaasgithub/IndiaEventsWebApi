using Aspose.Pdf.Plugins;
using IndiaEvents.Models.Models.MasterSheets;
using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Runtime.InteropServices;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]

    public class FCPAMasterController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        public FCPAMasterController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }






        [HttpGet("GetHCPDataUsingMISCode")]
        public IActionResult GetHCPDataUsingMISCode(string misCode)
        {

            string[] sheetIds = {
        //configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
        //configuration.GetSection("SmartsheetSettings:HcpMaster4").Value,
        configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value,
        configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value
    };

            foreach (string i in sheetIds)
            {
                long.TryParse(i, out long p);
                Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);

                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                Column fcpaSignOffDateColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "FCPA Sign Off Date");
                Column fcpaExpiryDateColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "FCPA Expiry Date");
                Column fcpaValidColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "FCPA Valid?");

                //if (misCodeColumn != null)
                //{
                //    Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                //        row.Cells != null &&
                //        row.Cells.Any(cell =>
                //            cell.ColumnId == misCodeColumn.Id &&
                //            cell.Value != null &&
                //            cell.Value.ToString() == misCode)
                //    );
                if (misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                        row.Cells != null &&

                        row.Cells.Any(cell =>
                            cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == misCode
                        )
                    );

                    if (existingRow != null)
                    {

                        Cell misCodeCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == misCodeColumn.Id);
                        Cell fcpaSignOffDateCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == fcpaSignOffDateColumn.Id);
                        Cell fcpaExpiryDateCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == fcpaExpiryDateColumn.Id);
                        Cell fcpaValidCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == fcpaValidColumn.Id);


                        var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(p, existingRow.Id.Value, null);
                        var url = "";
                        foreach (var attachment in attachments.Data)
                        {
                            if (attachment != null)
                            {
                                var AID = (long)attachment.Id;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(p, AID);
                                url = file.Url;

                            }
                        }

                        return Ok(new
                        {
                            MISCode = misCodeCell?.Value,
                            FCPASignOffDate = fcpaSignOffDateCell?.Value,
                            FCPAExpiryDate = fcpaExpiryDateCell?.Value,
                            FCPAValid = fcpaValidCell?.Value,
                            Attachments = attachments,
                            Url = url

                        });
                    }
                }
            }

            return Ok("False");
        }

        [HttpPut("UploadFcpaFileAndDate")]
        public IActionResult UploadFcpaFileAndDate(FcpaUpload formdata)
        {
            try
            {
                string SheetId1 = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
                string SheetId2 = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;
                //string SheetId3 = configuration.GetSection("SmartsheetSettings:HcpMaster").Value;
                string SheetId4 = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
                string SheetId5 = configuration.GetSection("SmartsheetSettings:HcpMaster2").Value;
                string SheetId6 = configuration.GetSection("SmartsheetSettings:HcpMaster3").Value;
                string SheetId7 = configuration.GetSection("SmartsheetSettings:HcpMaster4").Value;

                List<string> files = new List<string> { SheetId4, SheetId5, SheetId6, SheetId7 };

                if (formdata.SelectedType == "Speaker" || formdata.SelectedType == "Chair Person" || formdata.SelectedType == "Panelist" || formdata.SelectedType == "Moderator")
                {
                    Sheet sheet = SheetHelper.GetSheetById(smartsheet, SheetId2);
                    //Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
                    var targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.MisCode));
                    if (targetRow != null)
                    {
                        var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments((long)sheet.Id, targetRow.Id.Value, null);
                        if (a.Data.Count > 0)
                        {
                            foreach (var i in a.Data)
                            {
                                long id = i.Id.Value;
                                smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                                  (long)sheet.Id,           // sheetId
                                  id            // attachmentId
                                );
                            }
                        }


                        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet, "FCPA Sign Off Date");
                        var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formdata.FcpaDate };
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                        if (cellToUpdate != null) { cellToUpdate.Value = formdata.FcpaDate; }

                        var addedrow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                        var name = "FCPA";
                        var val = "";
                        var filePath = SheetHelper.testingFile(formdata.UploadFile, val, name);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                    sheet.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                    }

                }
                else if (formdata.SelectedType == "Trainer")
                {
                    Sheet sheet = SheetHelper.GetSheetById(smartsheet, SheetId1);
                    // Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
                    var targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.MisCode));

                    if (targetRow != null)
                    {
                        var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments((long)sheet.Id, targetRow.Id.Value, null);
                        if (a.Data.Count > 0)
                        {
                            foreach (var i in a.Data)
                            {
                                long id = i.Id.Value;
                                smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                                  (long)sheet.Id,           // sheetId
                                  id            // attachmentId
                                );
                            }
                        }


                        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet, "FCPA Sign Off Date");
                        var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formdata.FcpaDate };
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                        if (cellToUpdate != null) { cellToUpdate.Value = formdata.FcpaDate; }

                        var addedrow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                        var name = "FCPA";
                        var val = "";
                        var filePath = SheetHelper.testingFile(formdata.UploadFile, val, name);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                    sheet.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                    }

                }
                else if (formdata.SelectedType == "Other")
                {
                    foreach (var sheetid in files)
                    {
                        Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetid);
                        var targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.MisCode));
                        if (targetRow != null)
                        {
                            var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments((long)sheet.Id, targetRow.Id.Value, null);
                            if (a.Data.Count > 0)
                            {
                                foreach (var i in a.Data)
                                {
                                    long id = i.Id.Value;
                                    smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                                      (long)sheet.Id,           // sheetId
                                      id            // attachmentId
                                    );
                                }
                            }


                            long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet, "FCPA Sign Off Date");
                            var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formdata.FcpaDate };
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                            var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                            if (cellToUpdate != null) { cellToUpdate.Value = formdata.FcpaDate; }

                            var addedrow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                            var name = "FCPA";
                            var val = "";
                            var filePath = SheetHelper.testingFile(formdata.UploadFile, val, name);
                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                        sheet.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                        }
                    }
                }



                return Ok(new { Message = "FCPA Upsated Successfully" });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }

    }
}

