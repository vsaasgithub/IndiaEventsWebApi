using IndiaEvents.Models.Models.Draft;
using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Collections.Concurrent;
using System.Drawing.Printing;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers.Draft
{
    [Route("api/[controller]")]
    [ApiController]
    public class DraftController : ControllerBase
    {
        private readonly string accessToken;
        private readonly string sheetId;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly Sheet sheet;
        public DraftController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            sheetId = configuration.GetSection("SmartsheetSettings:DraftSheet").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
        }
        [HttpPost("PostDraftData")]
        public IActionResult PostDraftData(PostDraftData formDataList)
        {
            try
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventDate"), Value = formDataList.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InitiatorName"), Value = formDataList.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formDataList.InitiatorEmail });

                var addedRows = smartsheet.SheetResources.RowResources.AddRows((long)sheet.Id, new Row[] { newRow });
                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet, "Draft ID");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                if (formDataList.Isfile == "Yes")
                {
                    foreach (var p in formDataList.UploadFiles)
                    {
                        string[] words = p.Split(':');
                        var r = words[0];
                        var q = words[1];
                        var name = r.Split(".")[0];
                        var filePath = SheetHelper.testingFile(q, val, name);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)sheet.Id, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
                return Ok(new
                { Message = $"Row with id {val} created Successfully!" });
            }

            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


        [HttpGet("GetDraftDatausingId")]
        public IActionResult GetDraftDatausingId(string id)
        {
            Dictionary<string, object> DraftData = new Dictionary<string, object>();
            Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));
            if (existingRow != null)
            {

                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                for (int i = 0; i < columnNames.Count; i++)
                {

                    DraftData[columnNames[i]] = existingRow.Cells[i].Value;

                }
                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, existingRow.Id.Value, null);
               
                List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
                foreach (var attachment in attachments.Data)
                {
                    var AID = (long)attachment.Id;
                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
                    {
                        { "Name", file.Name.Split("-")[1] },
                        { "Url", file.Url }
                    };
                    attachmentsList.Add(attachmentInfo);
                }
                DraftData["Attachments"] = attachmentsList;
            }
            return Ok(DraftData);
        }

        [HttpDelete("DeleteRowUsingEventId")]
        public IActionResult DeleteRowUsingEventId(string id)
        {
            try
            {            
                Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));
                if (existingRow == null)
                {
                    return NotFound($"Row with id {id} not found.");
                }
                var Id = existingRow.Id.Value;
                smartsheet.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { Id }, true);

                return Ok(new { Message = "Data Deleted successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPut("UpdateDraftData")]
        public IActionResult UpdateDraftData(DraftData formDataList)
        {
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

            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
              
            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {

                string rowData = $" {formdata.MIS} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();

            foreach (var formdata in formDataList.RequestBrandsList)
            {

                string rowData = $"BrandName: {formdata.BrandName} |ProjectId: {formdata.ProjectId} |PercentAllocation: {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.EventRequestInvitees)
            {

                 string rowData = $"InviteeName: {formdata.InviteeName} | LocalConveyance: {formdata.LocalConveyance} | BtcorBte: {formdata.BtcorBte} | LcAmount: {formdata.LcAmount} | InviteedFrom: {formdata.InviteedFrom} | Speciality: {formdata.Speciality} | HCPType: {formdata.HCPType} ";
                //string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                addedInviteesData.AppendLine(rowData);
                addedInviteesDataNo++;
               
            }
            string Invitees = addedInviteesData.ToString();


            foreach (var formdata in formDataList.EventRequestHcpRole)
            {

              
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
               // string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {} |Trav.&Acc.Amt: {} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
              
            }
            string HCP = addedHcpData.ToString();
            return Ok();

        }


    }
}

