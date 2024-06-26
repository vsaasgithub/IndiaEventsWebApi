﻿using Aspose.Pdf.Plugins;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.MasterSheets.CodeCreation;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace IndiaEventsWebApi.Controllers.MasterSheets.CodeCreation
{
    [Route("api/[controller]")]
    [ApiController]
    public class TrainerCodeCreationController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public TrainerCodeCreationController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetTrainerCodeCreationData")]
        public IActionResult GetSpeakerCodeCreationData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:TrainerCodeCreation").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet.Rows)
                {
                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                    {
                        rowData[columnNames[i]] = row.Cells[i].Value;

                    }
                    sheetData.Add(rowData);
                }
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("AddTrainersData")]
        public IActionResult AddTrainersData(TrainerCodeGeneration formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:TrainerCodeCreation").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                string[] sheetIds = {

                configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value


                };
                var mis = "";
                var sheetval = "";
                string trainerType = formData.TrainerType;
                foreach (string i in sheetIds)
                {
                    long.TryParse(i, out long p);
                    Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);

                    Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");

                    if (misCodeColumn != null)
                    {
                        Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                        row.Cells != null &&

                        row.Cells.Any(cell =>
                            cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == formData.MisCode
                        )
                        );
                        if (existingRow != null)
                        {
                            mis = formData.MisCode;
                            sheetval = sheeti.Name;


                        }
                    }
                }
                if (mis != "" && mis.ToLower() != "na")
                {
                    return Ok(new
                    {
                        Message = $"MIS Code: {formData.MisCode} already exist in sheetname:{sheetval}"
                    });
                }

                else
                {
                    var IsTrainerCV = "";
                    var IsTrainerCertificate = "";

                    if (formData.TrainerCV != "")
                    {
                        IsTrainerCV = "Yes";
                    }
                    else
                    {
                        IsTrainerCV = "No";
                    }
                    if (formData.Trainercertificate != "")
                    {
                        IsTrainerCertificate = "Yes";
                    }
                    else
                    {
                        IsTrainerCertificate = "No";
                    }
                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.InitiatorNameName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "TrainerName"), Value = formData.TrainerName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Brand"), Value = formData.TrainerBrand });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MisCode"), Value = SheetHelper.MisCodeCheck(formData.MisCode) });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Division"), Value = formData.Division });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speciality"), Value = formData.Speciality });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Qualification"), Value = formData.Qualification });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Address"), Value = formData.Address });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Country"), Value = formData.Country });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Contact Number"), Value = formData.Contact_Number });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Type"), Value = formData.TrainerType });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trained by"), Value = formData.Trainedby });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer CV"), Value = IsTrainerCV });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer certificate"), Value = IsTrainerCertificate });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trained on"), Value = formData.Trainedon });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Category"), Value = formData.Trainer_Category });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Criteria"), Value = formData.Trainer_Criteria });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Criteria Details"), Value = formData.Trainer_Criteria_Details });

                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    var RowId = addedRows[0].Id.Value;


                    if (IsTrainerCV == "Yes")
                    {
                        //var addFile = AddFile(formData.TrainerCV,RowId );
                        byte[] fileBytes = Convert.FromBase64String(formData.TrainerCV);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = SheetHelper.GetFileType(fileBytes);
                        string fileName = " TrainerCV." + fileType;
                        // string fileName = val+x + ": AttachedFile." + fileType;
                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addedRows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        // string type = SheetHelper.GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId, addedRow.Id.Value, filePath, "application/msword");

                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }
                    if (IsTrainerCertificate == "Yes")
                    {
                        //var addFile = AddFile(formData.TrainerCV,RowId );
                        byte[] fileBytes = Convert.FromBase64String(formData.Trainercertificate);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = SheetHelper.GetFileType(fileBytes);
                        string fileName = " TrainerCertificate." + fileType;
                        // string fileName = val+x + ": AttachedFile." + fileType;
                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addedRows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        // string type = SheetHelper.GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId, addedRow.Id.Value, filePath, "application/msword");


                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }


                }
                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        [HttpPut("UpdateTrainerDatausingTrainerId")]
        public IActionResult UpdateVendorDatausingVendorId(UpdateTrainerCodeGeneration formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
                //string sheetId1 = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                long.TryParse(sheetId, out long parsedSheetId);


                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                var IsTrainerCV = !string.IsNullOrEmpty(formData.TrainerCV) ? "Yes" : "No";
                var IsTrainerCertificate = !string.IsNullOrEmpty(formData.Trainercertificate) ? "Yes" : "No";
                Row existingRow = GetRowById(smartsheet, parsedSheetId, formData.TrainerId);
                Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };

                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.InitiatorNameName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "TrainerName"), Value = formData.TrainerName });
                //newRow.Cells.Add(new Cell  { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Code"),  Value = formData.TrainerCode });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Brand"), Value = formData.TrainerBrand });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MisCode"), Value = SheetHelper.MisCodeCheck(formData.MisCode) });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Division"), Value = formData.Division });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speciality"), Value = formData.Speciality });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Qualification"), Value = formData.Qualification });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Address"), Value = formData.Address });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Country"), Value = formData.Country });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Contact Number"), Value = formData.Contact_Number });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Type"), Value = formData.TrainerType });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trained by"), Value = formData.Trainedby });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer CV"), Value = IsTrainerCV });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer certificate"), Value = IsTrainerCertificate });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trained on"), Value = formData.Trainedon });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Category"), Value = formData.Trainer_Category });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Criteria"), Value = formData.Trainer_Criteria });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Trainer Criteria Details"), Value = formData.Trainer_Criteria_Details });


                var updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });

                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(parsedSheetId, updatedRow[0].Id.Value, null);

                if (IsTrainerCV == "Yes")
                {
                    foreach (var a in attachments.Data)
                    {
                        long Id = (long)a.Id;
                        var Fullname = a.Name.Split(".");
                        var splitName = Fullname[0];

                        if (splitName == " TrainerCV")
                        {

                            smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                              parsedSheetId,           // sheetId
                              Id            // attachmentId
                            );

                        }
                    }


                    byte[] fileBytes = Convert.FromBase64String(formData.TrainerCV);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = SheetHelper.GetFileType(fileBytes);
                    string fileName = " ChequeDocument." + fileType;

                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = updatedRow[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    // string type = SheetHelper.GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId, addedRow.Id.Value, filePath, "application/msword");

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }

                if (IsTrainerCertificate == "Yes")
                {
                    foreach (var a in attachments.Data)
                    {
                        long Id = (long)a.Id;
                        var Fullname = a.Name.Split(".");
                        var splitName = Fullname[0];


                        if (splitName == " PanCardDocument")
                        {

                            smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                              parsedSheetId,           // sheetId
                              Id            // attachmentId
                            );

                        }
                    }

                    //var addFile = AddFile(formData.TrainerCV,RowId );
                    byte[] fileBytes = Convert.FromBase64String(formData.Trainercertificate);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = SheetHelper.GetFileType(fileBytes);
                    string fileName = " PanCardDocument." + fileType;
                    // string fileName = val+x + ": AttachedFile." + fileType;
                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = updatedRow[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    // string type = SheetHelper.GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId, addedRow.Id.Value, filePath, "application/msword");
                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }

                }

                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }

        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string email)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);



            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "TrainerId");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == email);

                    if (cell != null)
                    {
                        return row;
                    }
                }
            }

            return null;
        }

       

    }
}
