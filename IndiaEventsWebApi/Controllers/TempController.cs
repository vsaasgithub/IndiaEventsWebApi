﻿//using Microsoft.AspNetCore.Mvc;
//using Smartsheet.Api;
//using Smartsheet.Api.Models;
//using IndiaEventsWebApi.Models.EventTypeSheets;
//using System.Text;
//using System.Data;
//using iTextSharp.text;
//using iTextSharp.text.pdf;
//using IndiaEventsWebApi.Junk.Test;
//using Microsoft.Extensions.Logging;
//using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;
//using IndiaEventsWebApi.Helper;

//namespace IndiaEventsWebApi.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class TempController : ControllerBase
//    {
//        private readonly string accessToken;
//        private readonly IConfiguration configuration;
//        public TempController(IConfiguration configuration)
//        {
//            this.configuration = configuration;
//            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
//        }





//        [HttpPost("AddFormData")]
//        public async Task<IActionResult> AddFormData(IFormFile i)
//        {
//            try
//            {

//                var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//                var sheetId1 = 6857973674495876;
//                var rowId = 2502681132175236;
//                Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

//                var newRow = new Row();
//                newRow.Cells = new List<Cell>();

//                // Upload the file to Smartsheet
//                var file = i;
//                if (file != null && file.Length > 0)
//                {
//                    //smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                    //       sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
//                    Attachment attachment = smartsheet.SheetResources.AttachmentResources.AttachFile(sheetId1, file.FileName, "application/msword");
//                   // newRow.Cells.Add(new Cell.AddCellBuilder(sheet.Columns[0].Id, attachment.AttachmentId).Build());
//                }
//                //using (var memoryStream = new MemoryStream())
//                //{
//                //    await file.CopyToAsync(memoryStream);
//                //    newRow.Cells.Add(new Cell.AddCellBuilder(sheet.Columns[0].Id, memoryStream.ToArray()).Build());
//                //}
            
//                else
//                {
//                    return BadRequest("No file uploaded.");
//                }

//                // Add the new row to the sheet
//                smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });



//                //byte[] fileBytes = Convert.FromBase64String(i);

//                //if (fileBytes != null && fileBytes.Length > 0)
//                //{
//                //    var fileName = "FCPA" + fileBytes;//fileUploadModel.File.FileName;
//                //    var folderName = Path.Combine("Resources", "Images");
//                //    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//                //    var fullPath = Path.Combine(pathToSave, fileName);

//                //    if (!Directory.Exists(pathToSave))
//                //    {
//                //        Directory.CreateDirectory(pathToSave);
//                //    }


//                //    using (var fileStream = new FileStream(fullPath, FileMode.Create))
//                //    {

//                //        fileBytes.;
//                //    }
//                //    //string type = fileBytes.;//fileUploadModel.File.ContentType;
//                //    var addedRow = addedRows[0];
//                //    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                //        sheetId1, addedRow.Id.Value, fullPath, "");


//                //}




//                return Ok("Data added successfully.");
//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }

//        }




































//        [HttpPost("Basestring")]
//        public IActionResult Basestring(string[] data)
//        {
//            foreach (var x in data)
//            {
//                // string phrase = "The quick:brown fox jumps over the lazy dog.";
//                string[] words = x.Split(':');
//                var p = words[0];
//                var q = words[1];

//                byte[] fileBytes = Convert.FromBase64String(q);
//                var folderName = Path.Combine("Resources", "Images");
//                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//                if (!Directory.Exists(pathToSave))
//                {
//                    Directory.CreateDirectory(pathToSave);
//                }

//                string fileType = SheetHelper.GetFileType(fileBytes);
//                string fileName = p + fileType;
//                // string fileName = val+x + ": AttachedFile." + fileType;
//                string filePath = Path.Combine(pathToSave, fileName);

//                System.IO.File.WriteAllBytes(filePath, fileBytes);
//                string type = SheetHelper.GetContentType(fileType);

//                //var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                //parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
//                //j++;

//                if (System.IO.File.Exists(filePath))
//                {
//                    System.IO.File.Delete(filePath);
//                }

//            }
//            return Ok("Null");
//        }


















//        [HttpGet("GenerateSummaryPDF")]
//        public void GenerateSummaryPDF(string EventID)
//        {
//            try
//            {

//                var EventCode = "";
//                var EventName = "";
//                var EventDate = "";
//                var EventVenue = "";

//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//                string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
//                long.TryParse(sheetId_SpeakerCode, out long parsedSheetId_SpeakerCode);
//                Sheet sheet_SpeakerCode = smartsheet.SheetResources.GetSheet(parsedSheetId_SpeakerCode, null, null, null, null, null, null, null);

//                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
//                long.TryParse(sheetId, out long parsedSheetId);
//                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

//                string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
//                long.TryParse(sheetId1, out long parsedSheetId1);
//                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

//                string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
//                long.TryParse(processSheet, out long parsedProcessSheet);
//                Sheet processSheetData = smartsheet.SheetResources.GetSheet(parsedProcessSheet, null, null, null, null, null, null, null);

//                long rowId = 0;

//                Column processIdColumn = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
//                if (processIdColumn != null)
//                {
//                    // Find all rows with the specified speciality
//                    List<Row> targetRows = processSheetData.Rows
//                        .Where(row => row.Cells.Any(cell => cell.ColumnId == processIdColumn.Id && cell.Value.ToString() == EventID))
//                        .ToList();

//                    if (targetRows.Any())
//                    {
//                        var rowIds = targetRows.Select(row => row.Id).ToList();
//                        rowId = (long)rowIds[0];
//                    }
//                }


//                Column SpecialityColumn = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
//                Column targetColumn1 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Topic", StringComparison.OrdinalIgnoreCase));
//                Column targetColumn2 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventDate", StringComparison.OrdinalIgnoreCase));
//                Column targetColumn3 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "VenueName", StringComparison.OrdinalIgnoreCase));

//                if (SpecialityColumn != null)
//                {
//                    Row targetRow = sheet1.Rows
//                     .FirstOrDefault(row => row.Cells?.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value?.ToString() == EventID) == true);

//                    if (targetRow != null)
//                    {

//                        EventCode = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == SpecialityColumn.Id)?.Value?.ToString();
//                        EventName = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn1.Id)?.Value?.ToString();
//                        EventDate = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn2.Id)?.Value?.ToString();
//                        EventVenue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn3.Id)?.Value?.ToString();
//                    }
//                }

//                List<string> requiredColumns = new List<string> { "HCPName", "MISCode", "Speciality", "HCP Type" };
//                List<Column> selectedColumns = sheet_SpeakerCode.Columns
//                    .Where(column => requiredColumns.Contains(column.Title, StringComparer.OrdinalIgnoreCase)).ToList();
//                DataTable dtMai = new DataTable();
//                dtMai.Columns.Add("S.No", typeof(int));
//                foreach (Column column in selectedColumns)
//                {
//                    dtMai.Columns.Add(column.Title);
//                }
//                dtMai.Columns.Add("Sign");
//                int Sr_No = 1;
//                foreach (Row row in sheet_SpeakerCode.Rows)
//                {
//                    string eventId = row.Cells
//                        .FirstOrDefault(cell => sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
//                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
//                    {
//                        DataRow newRow = dtMai.NewRow();
//                        newRow["S.No"] = Sr_No;
//                        foreach (Cell cell in row.Cells)
//                        {
//                            string columnName = sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
//                            if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
//                            {
//                                newRow[columnName] = cell.DisplayValue;
//                            }
//                        }
//                        dtMai.Rows.Add(newRow);
//                        Sr_No++;
//                    }
//                }

//                foreach (Row row in sheet.Rows)
//                {
//                    string eventId = row.Cells.FirstOrDefault(cell => sheet.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
//                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
//                    {
//                        DataRow newRow = dtMai.NewRow();
//                        newRow["S.No"] = Sr_No;
//                        foreach (Cell cell in row.Cells)
//                        {
//                            string columnName = sheet.Columns
//                                .FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
//                            if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
//                            {
//                                newRow[columnName] = cell.DisplayValue;
//                            }
//                        }
//                        dtMai.Rows.Add(newRow);
//                        Sr_No++;
//                    }
//                }
//                byte[] fileBytes = exportpdf(dtMai, EventCode, EventName, EventDate, EventVenue, dtMai);
//                string filename = "Sample_PDF_" + EventID + ".pdf";
//                var folderName = Path.Combine("Resources", "Images");
//                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//                if (!Directory.Exists(pathToSave))
//                {
//                    Directory.CreateDirectory(pathToSave);
//                }
//                string fileType = SheetHelper.GetFileType(fileBytes);
//                string filePath = Path.Combine(pathToSave, filename);
//                System.IO.File.WriteAllBytes(filePath, fileBytes);
//                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedProcessSheet, rowId, filePath, "application/msword");
//                if (System.IO.File.Exists(filePath))
//                {
//                    System.IO.File.Delete(filePath);
//                }
//                List<Attachment> attachments = new List<Attachment>();

//                foreach (Row row in sheet_SpeakerCode.Rows)
//                {
//                    Cell matchingCell = row.Cells.FirstOrDefault(cell => cell.DisplayValue == EventID);

//                    if (matchingCell != null && matchingCell.Value != null)
//                    {
//                        var Id = (long)row.Id;
//                        string eventId = matchingCell.Value.ToString();
//                        if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
//                        {
//                            var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(parsedSheetId_SpeakerCode, Id, null);
//                            var url = "";
//                            foreach (var x in a.Data)
//                            {
//                                if (x != null)
//                                {
//                                    var AID = (long)x.Id;
//                                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(parsedSheetId_SpeakerCode, AID);
//                                    url = file.Url;
//                                }
//                                if (url != "")
//                                {
//                                    using (HttpClient client = new HttpClient())
//                                    {
//                                        byte[] data = client.GetByteArrayAsync(url).Result;
//                                        string base64 = Convert.ToBase64String(data);
//                                        byte[] xy = Convert.FromBase64String(base64);
//                                        var f = Path.Combine("Resources", "Images");
//                                        var ps = Path.Combine(Directory.GetCurrentDirectory(), f);
//                                        if (!Directory.Exists(ps))
//                                        {
//                                            Directory.CreateDirectory(ps);
//                                        }
//                                        string ft = SheetHelper.GetFileType(xy);
//                                        string fileName = eventId + "-" + x + " AttachedFile." + ft;
//                                        string fp = Path.Combine(ps, fileName);
//                                        var addedRow = rowId;
//                                        System.IO.File.WriteAllBytes(fp, xy);
//                                        string type = SheetHelper.GetContentType(ft);
//                                        var z = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedProcessSheet, addedRow, fp, "application/msword");
//                                    }
//                                    var bs64 = "";
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                //return BadRequest(ex.Message);
//            }
//        }




//        // Testing using row id
//        [HttpGet("GenerateSummaryTest")]
//        public IActionResult GenerateSummaryTest(long rowId)
//        {
//            try
//            {
//                var EventID = "";


//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();



//                string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
//                long.TryParse(processSheet, out long parsedProcessSheet);
//                Sheet processSheetData = smartsheet.SheetResources.GetSheet(parsedProcessSheet, null, null, null, null, null, null, null);

//                Column processIdColumn = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));

//                Row targetRow = processSheetData.Rows.FirstOrDefault(row => row.Id == rowId);


//                if (targetRow != null)
//                {



//                    if (processIdColumn != null)
//                    {
//                        var columnValue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn.Id)?.Value;
//                        EventID = columnValue?.ToString();
//                    }

//                }
//                else
//                {

//                    return BadRequest("Row not found in the specified sheet.");
//                }



//            }
//            catch (Exception ex)
//            {
//                //return BadRequest(ex.Message);
//            }
//            return Ok("created and attached in process sheet ...");
//        }






//        // testing sample pdf without passing eventid for webhook part

//        [HttpGet("GenerateSummaryPDFTest")]
//        public IActionResult GenerateSummaryPDFTest()
//        {
//            try
//            {
//                var EventID = "RQID319";
//                var EventCode = "";
//                var EventName = "";
//                var EventDate = "";
//                var EventVenue = "";

//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//                string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
//                long.TryParse(sheetId_SpeakerCode, out long parsedSheetId_SpeakerCode);
//                Sheet sheet_SpeakerCode = smartsheet.SheetResources.GetSheet(parsedSheetId_SpeakerCode, null, null, null, null, null, null, null);

//                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
//                long.TryParse(sheetId, out long parsedSheetId);
//                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

//                string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
//                long.TryParse(sheetId1, out long parsedSheetId1);
//                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

//                string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
//                long.TryParse(processSheet, out long parsedProcessSheet);
//                Sheet processSheetData = smartsheet.SheetResources.GetSheet(parsedProcessSheet, null, null, null, null, null, null, null);

//                long rowId = 0;

//                Column processIdColumn = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
//                if (processIdColumn != null)
//                {
//                    // Find all rows with the specified speciality
//                    List<Row> targetRows = processSheetData.Rows
//                        .Where(row => row.Cells.Any(cell => cell.ColumnId == processIdColumn.Id && cell.Value.ToString() == EventID))
//                        .ToList();

//                    if (targetRows.Any())
//                    {
//                        var rowIds = targetRows.Select(row => row.Id).ToList();
//                        rowId = (long)rowIds[0];
//                    }
//                }


//                Column SpecialityColumn = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
//                Column targetColumn1 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Topic", StringComparison.OrdinalIgnoreCase));
//                Column targetColumn2 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventDate", StringComparison.OrdinalIgnoreCase));
//                Column targetColumn3 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "VenueName", StringComparison.OrdinalIgnoreCase));

//                if (SpecialityColumn != null)
//                {
//                    Row targetRow = sheet1.Rows
//                     .FirstOrDefault(row => row.Cells?.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value?.ToString() == EventID) == true);

//                    if (targetRow != null)
//                    {

//                        EventCode = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == SpecialityColumn.Id)?.Value?.ToString();
//                        EventName = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn1.Id)?.Value?.ToString();
//                        EventDate = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn2.Id)?.Value?.ToString();
//                        EventVenue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn3.Id)?.Value?.ToString();
//                    }
//                }

//                List<string> requiredColumns = new List<string> { "HCPName", "MISCode", "Speciality", "HCP Type" };
//                List<Column> selectedColumns = sheet_SpeakerCode.Columns
//                    .Where(column => requiredColumns.Contains(column.Title, StringComparer.OrdinalIgnoreCase)).ToList();
//                DataTable dtMai = new DataTable();
//                dtMai.Columns.Add("S.No", typeof(int));
//                foreach (Column column in selectedColumns)
//                {
//                    dtMai.Columns.Add(column.Title);
//                }
//                dtMai.Columns.Add("Sign");
//                int Sr_No = 1;
//                foreach (Row row in sheet_SpeakerCode.Rows)
//                {
//                    string eventId = row.Cells
//                        .FirstOrDefault(cell => sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
//                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
//                    {
//                        DataRow newRow = dtMai.NewRow();
//                        newRow["S.No"] = Sr_No;
//                        foreach (Cell cell in row.Cells)
//                        {
//                            string columnName = sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
//                            if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
//                            {
//                                newRow[columnName] = cell.DisplayValue;
//                            }
//                        }
//                        dtMai.Rows.Add(newRow);
//                        Sr_No++;
//                    }
//                }

//                foreach (Row row in sheet.Rows)
//                {
//                    string eventId = row.Cells.FirstOrDefault(cell => sheet.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
//                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
//                    {
//                        DataRow newRow = dtMai.NewRow();
//                        newRow["S.No"] = Sr_No;
//                        foreach (Cell cell in row.Cells)
//                        {
//                            string columnName = sheet.Columns
//                                .FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
//                            if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
//                            {
//                                newRow[columnName] = cell.DisplayValue;
//                            }
//                        }
//                        dtMai.Rows.Add(newRow);
//                        Sr_No++;
//                    }
//                }
//                byte[] fileBytes = exportpdf(dtMai, EventCode, EventName, EventDate, EventVenue, dtMai);
//                string filename = "Sample_PDF_" + EventID + ".pdf";
//                var folderName = Path.Combine("Resources", "Images");
//                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//                if (!Directory.Exists(pathToSave))
//                {
//                    Directory.CreateDirectory(pathToSave);
//                }
//                string fileType = SheetHelper.GetFileType(fileBytes);
//                string filePath = Path.Combine(pathToSave, filename);
//                System.IO.File.WriteAllBytes(filePath, fileBytes);
//                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedProcessSheet, rowId, filePath, "application/msword");
//                if (System.IO.File.Exists(filePath))
//                {
//                    System.IO.File.Delete(filePath);
//                }
//                List<Attachment> attachments = new List<Attachment>();

//                foreach (Row row in sheet_SpeakerCode.Rows)
//                {
//                    Cell matchingCell = row.Cells.FirstOrDefault(cell => cell.DisplayValue == EventID);

//                    if (matchingCell != null && matchingCell.Value != null)
//                    {
//                        var Id = (long)row.Id;
//                        string eventId = matchingCell.Value.ToString();
//                        if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
//                        {
//                            var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(parsedSheetId_SpeakerCode, Id, null);
//                            var url = "";
//                            foreach (var x in a.Data)
//                            {
//                                if (x != null)
//                                {
//                                    var AID = (long)x.Id;
//                                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(parsedSheetId_SpeakerCode, AID);
//                                    url = file.Url;
//                                }
//                                if (url != "")
//                                {
//                                    using (HttpClient client = new HttpClient())
//                                    {
//                                        byte[] data = client.GetByteArrayAsync(url).Result;
//                                        string base64 = Convert.ToBase64String(data);
//                                        byte[] xy = Convert.FromBase64String(base64);
//                                        var f = Path.Combine("Resources", "Images");
//                                        var ps = Path.Combine(Directory.GetCurrentDirectory(), f);
//                                        if (!Directory.Exists(ps))
//                                        {
//                                            Directory.CreateDirectory(ps);
//                                        }
//                                        string ft = SheetHelper.GetFileType(xy);
//                                        string fileName = eventId + "-" + x + " AttachedFile." + ft;
//                                        string fp = Path.Combine(ps, fileName);
//                                        var addedRow = rowId;
//                                        System.IO.File.WriteAllBytes(fp, xy);
//                                        string type = SheetHelper.GetContentType(ft);
//                                        var z = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedProcessSheet, addedRow, fp, "application/msword");
//                                    }
//                                    var bs64 = "";
//                                }
//                            }
//                        }
//                    }
//                }


//            }
//            catch (Exception ex)
//            {
//                //return BadRequest(ex.Message);
//            }
//            return Ok("created and attached in process sheet ...");
//        }





//        [HttpPost("GeneratePdfAndAttachments")]
//        public async Task<IActionResult> GeneratePdfAndAttachments(string EventId)
//        {
//            try
//            {
//                int timeInterval = 300000;
//                await Task.Delay(timeInterval);

//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//                string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
//                long.TryParse(sheetId_SpeakerCode, out long parsedSheetId_SpeakerCode);
//                Sheet sheet_SpeakerCode = smartsheet.SheetResources.GetSheet(parsedSheetId_SpeakerCode, null, null, null, null, null, null, null);


//                string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
//                long.TryParse(processSheet, out long parsedProcessSheet);
//                Sheet processSheetData = smartsheet.SheetResources.GetSheet(parsedProcessSheet, null, null, null, null, null, null, null);
//                long rowId = 0;

//                Column processIdColumn = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
//                if (processIdColumn != null)
//                {
//                    // Find all rows with the specified speciality
//                    List<Row> targetRows = processSheetData.Rows
//                        .Where(row => row.Cells.Any(cell => cell.ColumnId == processIdColumn.Id && cell.Value.ToString() == EventId))
//                        .ToList();

//                    if (targetRows.Any())
//                    {
//                        var rowIds = targetRows.Select(row => row.Id).ToList();
//                        rowId = (long)rowIds[0];
//                    }

//                }

//                List<Attachment> attachments = new List<Attachment>();

//                foreach (Row row in sheet_SpeakerCode.Rows)
//                {
//                    Cell matchingCell = row.Cells.FirstOrDefault(cell => cell.DisplayValue == EventId);

//                    if (matchingCell != null && matchingCell.Value != null)
//                    {
//                        var Id = (long)row.Id;
//                        string eventId = matchingCell.Value.ToString();
//                        if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventId, StringComparison.OrdinalIgnoreCase))
//                        {
//                            var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(parsedSheetId_SpeakerCode, Id, null);
//                            var url = "";
//                            foreach (var x in a.Data)
//                            {
//                                if (x != null)
//                                {
//                                    var AID = (long)x.Id;
//                                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(parsedSheetId_SpeakerCode, AID);
//                                    url = file.Url;
//                                }
//                                if (url != "")
//                                {
//                                    using (HttpClient client = new HttpClient())
//                                    {
//                                        byte[] data = client.GetByteArrayAsync(url).Result;
//                                        string base64 = Convert.ToBase64String(data);
//                                        byte[] xy = Convert.FromBase64String(base64);
//                                        var f = Path.Combine("Resources", "Images");
//                                        var ps = Path.Combine(Directory.GetCurrentDirectory(), f);
//                                        if (!Directory.Exists(ps))
//                                        {
//                                            Directory.CreateDirectory(ps);
//                                        }
//                                        string ft = SheetHelper.GetFileType(xy);
//                                        string fileName = eventId + "-" + x + " AttachedFile." + ft;
//                                        string fp = Path.Combine(ps, fileName);
//                                        var addedRow = rowId;
//                                        System.IO.File.WriteAllBytes(fp, xy);
//                                        string type = SheetHelper.GetContentType(ft);
//                                        var z = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedProcessSheet, addedRow, fp, "application/msword");
//                                    }

//                                }
//                            }
//                        }
//                    }
//                }
//                return Ok("Done");

//            }




//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }
//        }


//        [HttpGet("GetcolumnsbysheetId")]
//        public IActionResult GetcolumnsbysheetId(long Id)
//        {
//            try
//            {
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//                //Sheet sheet = smartsheet.SheetResources.GetSheet(Id, null, null, null, null, null, null, null);
//                PaginatedResult<Column> columns = smartsheet.SheetResources.ColumnResources.ListColumns(
//               Id,               // sheetId
//               null,                           // IEnumerable<ColumnInclusion> include
//               null,                           // PaginationParameters
//               2                               // int compatibilityLevel
//             );

//                List<string> columnNames = new List<string>();
//                foreach (Column column in columns.Data)
//                {

//                    var cn = column.Title;
//                    var id = column.Id;
//                    var str = $"{cn} : {id}";
//                    columnNames.Add(str);

//                }

//                return Ok(columnNames);


//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }



//        }



//        [HttpGet("webhookhandler")]
//        public IActionResult webhookhandler()
//        {
//            try
//            {
//                var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//                var sheetId = 2789422625935236;
//                var webhookId = "";
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//                Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

//                Webhook smartsheetWebhook = smartsheet.WebhookResources.GetWebhook(sheetId);
//                //IndiaEventsWebApi.Controllers.WebHooksController.Webhook customWebhook = new IndiaEventsWebApi.Controllers.WebHooksController.Webhook
//                //{
//                //    // Assign properties from Smartsheet Webhook to your custom Webhook
//                //    // For example:
//                //    Id = smartsheetWebhook.Id,
//                //    Name = smartsheetWebhook.Name,
//                //    // Add other properties as needed
//                //};
//                //Webhook Webhook = smartsheet.WebhookResources.GetWebhook(sheetId);
//                return Ok();

//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.StackTrace);
//            }
//        }

//        private static DataTable ToDictionary(List<Dictionary<string, object>> list)
//        {
//            DataTable result = new DataTable();
//            if (list.Count == 0)
//                return result;

//            result.Columns.AddRange(
//                list.First().Select(r => new DataColumn(r.Key)).ToArray()
//            );

//            list.ForEach(r => result.Rows.Add(r.Select(c => c.Value).Cast<object>().ToArray()));

//            return result;
//        }

//        private byte[] exportpdf(DataTable dtEmployee, string EventCode, string EventName, string EventDate, string EventVenue, DataTable dtMai)
//        {
//            System.IO.MemoryStream ms = new System.IO.MemoryStream();
//            iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(PageSize.A4);
//            rec.BackgroundColor = new BaseColor(System.Drawing.Color.Olive);
//            Document doc = new Document(rec);
//            doc.SetPageSize(iTextSharp.text.PageSize.A4);
//            PdfWriter writer = PdfWriter.GetInstance(doc, ms);
//            doc.Open();
//            BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
//            iTextSharp.text.Font fntHead = new iTextSharp.text.Font(bfntHead, 16, 1, iTextSharp.text.BaseColor.BLUE);
//            Paragraph prgHeading = new Paragraph();
//            prgHeading.Alignment = Element.ALIGN_LEFT;
//            prgHeading.Add(new Chunk("Attendance Sheet".ToUpper(), fntHead));
//            doc.Add(prgHeading);
//            Paragraph prgGeneratedBY = new Paragraph();
//            BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
//            iTextSharp.text.Font fntAuthor = new iTextSharp.text.Font(btnAuthor, 8, 2, iTextSharp.text.BaseColor.BLUE);
//            prgGeneratedBY.Alignment = Element.ALIGN_RIGHT;
//            doc.Add(prgGeneratedBY);
//            Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, iTextSharp.text.BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
//            doc.Add(p);
//            doc.Add(new Chunk("\n", fntHead));

//            Paragraph pBody = new Paragraph();
//            pBody.Add(new Chunk("Event Code:" + EventCode));
//            pBody.Add(new Chunk("\nEvent Name:" + EventName));
//            pBody.Add(new Chunk("\nEvent Date:" + EventDate));
//            pBody.Add(new Chunk("\nEvent Venue:" + EventVenue));
//            pBody.Add(new Chunk("\n\nSpeakers: "));

//            //foreach(DataRow row in dtMai.Rows)
//            //{
//            //    string hcpName = row["HCPName"].ToString();
//            //    pBody.Add(new Chunk(" " + hcpName));
//            //}

//            string hcpNames = string.Join(", ", dtMai.AsEnumerable().Select(row => row["HCPName"].ToString()));
//            pBody.Add(new Chunk(" " + hcpNames));
//            doc.Add(pBody);
//            doc.Add(new Paragraph("\n "));
//            PdfPTable table = new PdfPTable(dtEmployee.Columns.Count);
//            table.WidthPercentage = 100;
//            float[] columnWidths = Enumerable.Range(0, dtEmployee.Columns.Count).Select(i => i == dtEmployee.Columns.IndexOf("HCPName") ? 2f : 1f).ToArray(); /*Count).Select(i => 1f).ToArray();*/
//            table.SetWidths(columnWidths);

//            for (int i = 0; i < dtEmployee.Columns.Count; i++)
//            {
//                string cellText = dtEmployee.Columns[i].ColumnName;
//                PdfPCell cell = new PdfPCell();
//                cell.Phrase = new Phrase(cellText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
//                cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
//                cell.HorizontalAlignment = Element.ALIGN_CENTER;
//                cell.PaddingBottom = 5;
//                table.AddCell(cell);
//            }
//            for (int i = 0; i < dtEmployee.Rows.Count; i++)
//            {
//                for (int j = 0; j < dtEmployee.Columns.Count; j++)
//                {
//                    table.AddCell(dtEmployee.Rows[i][j].ToString());
//                }
//            }
//            doc.Add(table);
//            doc.Close();
//            byte[] result = ms.ToArray();
//            return result;
//        }

//        private string GetContentType(string fileExtension)
//        {
//            switch (fileExtension.ToLower())
//            {
//                case "jpg":
//                case "jpeg":
//                    return "image/jpeg";
//                case "pdf":
//                    return "application/pdf";
//                case "gif":
//                    return "image/gif";
//                case "png":
//                    return "image/png";
//                case "webp":
//                    return "image/webp";
//                case "doc":
//                    return "application/msword";
//                case "docx":
//                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
//                default:
//                    return "application/octet-stream";
//            }
//        }

//        private string GetFileType(byte[] bytes)
//        {

//            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8)
//            {
//                return "jpg";
//            }
//            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "%PDF")
//            {
//                return "pdf";
//            }
//            else if (bytes.Length >= 3 && Encoding.UTF8.GetString(bytes, 0, 3) == "GIF")
//            {
//                return "gif";
//            }
//            else if (bytes.Length >= 8 && Encoding.UTF8.GetString(bytes, 0, 8) == "PNG\r\n\x1A\n")
//            {
//                return "png";
//            }
//            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "RIFF" && Encoding.UTF8.GetString(bytes, 8, 4) == "WEBP")
//            {
//                return "webp";
//            }
//            else if (bytes.Length >= 4 && (bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0))
//            {
//                return "doc";
//            }
//            else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
//            {
//                return "docx";
//            }
//            else
//            {
//                return "unknown";
//            }
//        }

//        private long GetColumnIdByName(Sheet sheet, string columnname)
//        {
//            foreach (var column in sheet.Columns)
//            {
//                if (column.Title == columnname)
//                {
//                    return column.Id.Value;
//                }
//            }
//            return 0;
//        }







//    }
//}





























































































////private List<Attachment> GetAttachmentsFromSheet(Sheet sheet, string eventID ,long sheetId)
////{
////    List<Attachment> attachments = new List<Attachment>();


////    foreach (Row row in sheet.Rows)
////    {
////        Cell matchingCell = row.Cells.FirstOrDefault(cell => cell.DisplayValue == eventID);

////        if (matchingCell != null && matchingCell.Value != null)
////        {
////            var rowId = row.Id;
////            string eventId = matchingCell.Value.ToString();

////            if (!string.IsNullOrEmpty(eventId) && eventId.Equals(eventID, StringComparison.OrdinalIgnoreCase))

////            {
////                var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheetId, rowId, null);
////                if (row.Attachments != null)
////                {
////                    attachments.AddRange(row.Attachments);
////                }

////            }
////        }
////    }

////    return attachments;
////}
////private getIds(Sheet sheet, string eventId)
////{
////    Column processIdColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
////    if (processIdColumn != null)
////    {
////        // Find all rows with the specified speciality
////        List<Row> targetRows = sheet.Rows
////            .Where(row => row.Cells.Any(cell => cell.ColumnId == processIdColumn.Id && cell.Value.ToString() == eventId))
////            .ToList();

////        if (targetRows.Any())
////        {
////            var rowIds = targetRows.Select(row => row.Id).ToList();
////            var rowId = (long)rowIds[0];
////            ;
////        }

////    }
////}






////[HttpPost("AddFile")]
////public IActionResult AddFormData(Class11 i)
////{
////    var file = i.File;
////    byte[] fileBytes = Convert.FromBase64String(i.File);


////}

////[HttpPost("AddFormData")]
////public IActionResult AddFormData(Class11 i)
////{
////    try
////    {

////        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
////        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

////        var sheetId1 = 6857973674495876;
////        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

////        var newRow = new Row();
////        newRow.Cells = new List<Cell>();

////        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

////        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });

////        byte[] fileBytes = Convert.FromBase64String(i.File);

////        var inputString = Encoding.UTF8.GetString(fileBytes);
////        //using (var ms = new MemoryStream(fileBytes, 0, fileBytes.Length))
////        //{
////        //    Image image = Image.FromStream(ms, true);
////        //    return image;
////        //}


////        if (inputString != null && inputString.Length > 0)
////        {
////            var fileName = "FCPA" + inputString;//fileUploadModel.File.FileName;
////            var folderName = Path.Combine("Resources", "Images");
////            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
////            var fullPath = Path.Combine(pathToSave, fileName);
////             System.IO.File.WriteAllBytes(fullPath, fileBytes);
////            if (!Directory.Exists(pathToSave))
////            {
////                Directory.CreateDirectory(pathToSave);
////            }


////            using (var fileStream = new FileStream(fullPath, FileMode.Create))
////            {

////                //fileBytes.CopyTo(FileStream);
////            }
////            //string type = fileBytes;//fileUploadModel.File.ContentType;
////            var addedRow = addedRows[0];
////            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
////                sheetId1, addedRow.Id.Value, fullPath, "application/msword");


////        }




////        return Ok("Data added successfully.");
////    }
////    catch (Exception ex)
////    {
////        return BadRequest(ex.Message);
////    }

////}





////NewMultipleUpload
////[HttpPost("AddFormData")]
////public IActionResult AddFormData(IFormFile[] fileUploadModel)
////{
////    try
////    {
////        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
////        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

////        var sheetId1 = 6857973674495876;
////        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

////        var newRow = new Row();
////        newRow.Cells = new List<Cell>();

////        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

////        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });

////        foreach (var k in fileUploadModel)
////        {
////            if (k != null && k.Length > 0)
////            {
////                var fileName = "FCPA" + k;//fileUploadModel.File.FileName;
////                var folderName = Path.Combine("Resources", "Images");
////                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
////                var fullPath = Path.Combine(pathToSave, fileName);

////                if (!Directory.Exists(pathToSave))
////                {
////                    Directory.CreateDirectory(pathToSave);
////                }


////                using (var fileStream = new FileStream(fullPath, FileMode.Create))
////                {
////                    k.CopyTo(fileStream);
////                }
////                string type = k.ContentType;//fileUploadModel.File.ContentType;
////                var addedRow = addedRows[0];
////                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
////                    sheetId1, addedRow.Id.Value, fullPath, type);


////            }
////        }



////        return Ok("Data added successfully.");
////    }
////    catch (Exception ex)
////    {
////        return BadRequest(ex.Message);
////    }

////}











////        //Single Upload
////        [HttpPost("AddFormData")]
////        public IActionResult AddFormData([FromForm] FileUploadodel fileUploadModel)
////        {

////            try
////            {
////                var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
////                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

////                var sheetId1 = 6857973674495876;
////                Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

////                var newRow = new Row();
////                newRow.Cells = new List<Cell>();

////                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

////                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });


////                if (fileUploadModel.File != null && fileUploadModel.File.Length > 0)
////                {

////                    var fileType = Path.GetExtension(fileUploadModel.File.FileName);
////                    var filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory());
////                    var docName = Path.GetFileName(fileUploadModel.File.FileName);
////                    Guid ID = Guid.NewGuid();
////                    //ExpenseDataRow.Guid = Convert.ToString(ID);
////                    //ExpenseDataRow.DocumentName = docName;
////                    //ExpenseDataRow.DocType = fileType;
////                    var docUrl = Path.Combine(filePath, "Resources\\Images", ID.ToString() + fileType);
////                    //ExpenseDataRow.DocUrl = Path.Combine(filePath, "assets\\Uploads", ID.ToString() + fileType);
////                    //using (var stream = new FileStream(ExpenseDataRow.DocUrl, FileMode.Create))
////                    using (var stream = System.IO.File.Create(docUrl))
////                    {
////                        //Attach_files[k].CopyToAsync(stream);
////                        fileUploadModel.File.CopyTo(stream);

////                    }








////                    //var fileName = "FCPA" + fileUploadModel.File.FileName;//fileUploadModel.File.FileName;
////                    //var folderName = Path.Combine("Resources", "Images");
////                    //var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
////                    //var fullPath = Path.Combine(pathToSave, fileName);

////                    //if (!Directory.Exists(pathToSave))
////                    //{
////                    //    Directory.CreateDirectory(pathToSave);
////                    //}
////                    //using (var stream = System.IO.File.Create(fullPath))
////                    //{
////                    //    //Attach_files[k].CopyToAsync(stream);
////                    //    fileUploadModel.File.CopyTo(stream);

////                    //}


////                    //using (var fileStream = new FileStream(fullPath, FileMode.Create))
////                    //{
////                    //    fileUploadModel.File.CopyTo(fileStream);
////                    //}
////                    string type = fileUploadModel.File.ContentType;//fileUploadModel.File.ContentType;
////                    var addedRow = addedRows[0];
////                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
////                        sheetId1, addedRow.Id.Value, docUrl, type /*"application/msword"*/);


////                }



////                return Ok("Data added successfully.");
////            }
////            catch (Exception ex)
////            {
////                return BadRequest(ex.Message);
////            }

////        }





////        //formultiple upload
////        //[HttpPost("AddFormData")]
////        //public IActionResult AddFormData( FileUploadodel fileUploadModel)
////        //{
////        //    try
////        //    {
////        //        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
////        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

////        //        var sheetId1 = 6857973674495876;
////        //        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

////        //        var newRow = new Row();
////        //        newRow.Cells = new List<Cell>();

////        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

////        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });
////        //        foreach (var p in fileUploadModel.File)
////        //        {
////        //            if (p != null && p.Length > 0)
////        //            {
////        //                var fileName = "FCPA"+p.FileName;//fileUploadModel.File.FileName;
////        //                var folderName = Path.Combine("Resources", "Images");
////        //                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
////        //                var fullPath = Path.Combine(pathToSave, fileName);

////        //                if (!Directory.Exists(pathToSave))
////        //                {
////        //                    Directory.CreateDirectory(pathToSave);
////        //                }


////        //                using (var fileStream = new FileStream(fullPath, FileMode.Create))
////        //                {
////        //                    p.CopyTo(fileStream);
////        //                }
////        //                string type = p.ContentType;//fileUploadModel.File.ContentType;
////        //                var addedRow = addedRows[0];
////        //                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
////        //                    sheetId1, addedRow.Id.Value, fullPath, type);


////        //            }
////        //        }


////        //        return Ok("Data added successfully.");
////        //    }
////        //    catch (Exception ex)
////        //    {
////        //        return BadRequest(ex.Message);
////        //    }

////        //}


////        //[HttpGet("GetFormDataByGender")]
////        //public IActionResult GetFormDataByGender(string gender)
////        //{
////        //    try
////        //    {
////        //        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
////        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
////        //        var sheetId1 = 6857973674495876;

////        //        
////        //        var genderColumnId = SheetHelper.GetColumnIdByName(smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null), "GENDER");

////        //        var filter = new SearchFilter
////        //        {
////        //            ColumnIds = new long[] { genderColumnId },
////        //            Text = gender
////        //        };

////        //        var filteredRows = smartsheet.SearchResources.Filter(sheetId1, filter);



////        //        return Ok(filteredRows);
////        //    }
////        //    catch (Exception ex)
////        //    {
////        //        return BadRequest(ex.Message);
////        //    }
////        //}





////        private long SheetHelper.GetColumnIdByName(Sheet sheet, string columnname)
////        {
////            foreach (var column in sheet.Columns)
////            {
////                if (column.Title == columnname)
////                {
////                    return column.Id.Value;
////                }
////            }
////            return 0;
////        }
////    }
////}

//////}
//////}

























//// ////////////////////////////////////////////////SINGLE FILE UPLOAD


////[HttpPost("AddFile")]
////public IActionResult AddFormData(Class11 i)
////{
////    try
////    {


////        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
////        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

////        var sheetId1 = 6857973674495876;
////        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

////        var newRow = new Row();
////        newRow.Cells = new List<Cell>();

////        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });


////        byte[] fileBytes = Convert.FromBase64String(i.File);
////        var folderName = Path.Combine("Resources", "Images");
////        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
////        if (!Directory.Exists(pathToSave))
////        {
////            Directory.CreateDirectory(pathToSave);
////        }

////        string fileType = SheetHelper.GetFileType(fileBytes);
////        string fileName = "ConvertedFile." + fileType;
////        string filePath = Path.Combine(pathToSave, fileName);

////        //string type = k.ContentType;//fileUploadModel.File.ContentType;
////        var addedRow = addedRows[0];

////        System.IO.File.WriteAllBytes(filePath, fileBytes);
////        string type = SheetHelper.GetContentType(fileType);
////        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
////                sheetId1, addedRow.Id.Value, filePath, "application/msword");




////        return Ok(new { Message = "Conversion successful", FilePath = filePath });


////    }
////    catch (Exception ex)
////    {
////        return BadRequest($"Error: {ex.Message}");
////    }
////}




