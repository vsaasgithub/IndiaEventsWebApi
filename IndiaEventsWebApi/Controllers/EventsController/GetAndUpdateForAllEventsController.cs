using IndiaEvents.Models.Models.Draft;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;

namespace IndiaEventsWebApi.Controllers.EventsController
{
    [Route("api/[controller]")]
    [ApiController]
    public class GetAndUpdateForAllEventsController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly string sheetId1;
        private readonly string sheetId2;
        private readonly string sheetId3;
        private readonly string sheetId4;
        private readonly string sheetId5;
        private readonly string sheetId6;
        private readonly string sheetId7;
        private readonly string sheetId8;
        private readonly string sheetId9;
        public GetAndUpdateForAllEventsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
            sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            sheetId8 = configuration.GetSection("SmartsheetSettings:EventRequestBeneficiary").Value;
            sheetId9 = configuration.GetSection("SmartsheetSettings:EventRequestProductBrandsList").Value;
        }
        #region
        //[HttpGet("GetDataFromAllSheetsUsingEventId")]
        //public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        //{
        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    //Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    //Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    //Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    //Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
        //    //Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
        //    //Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, sheetId9);
        //    Dictionary<string, object> DraftData = new();
        //    Row existingRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));

        //    if (existingRow != null)
        //    {
        //        List<string> columnNames = new List<string>();
        //        foreach (Column column in sheet1.Columns)
        //        {
        //            columnNames.Add(column.Title);
        //        }
        //        for (int i = 0; i < columnNames.Count; i++)
        //        {
        //            var x = existingRow.Cells[i].Value;
        //            if (columnNames[i] == "Brands")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> brandsList = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = brandsList;
        //                }

        //            }
        //            else if (columnNames[i] == "Panelists")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Panelists = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Panelists;
        //                }
        //            }
        //            else if (columnNames[i] == "SlideKits")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> SlideKits = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = SlideKits;
        //                }
        //            }
        //            else if (columnNames[i] == "Invitees")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Invitees = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Invitees;
        //                }
        //            }


        //            else if (columnNames[i] == "Expenses")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Expenses = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Expenses;

        //                }
        //            }
        //            else
        //            {
        //                DraftData[columnNames[i]] = existingRow.Cells[i].Value;
        //            }


        //        }
        //        var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, existingRow.Id.Value, null);

        //        List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();

        //        foreach (var attachment in attachments.Data)
        //        {
        //            var AID = (long)attachment.Id;
        //            var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);
        //            Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
        //            {
        //                { "Name", file.Name },
        //                { "Url", file.Url }
        //            };
        //            attachmentsList.Add(attachmentInfo);
        //        }
        //        DraftData["Attachments"] = attachmentsList;
        //    }
        //    return Ok(DraftData);





        //}
        #endregion

        //[HttpGet("GetDataFromAllSheetsUsingEventIdInPreEvent")]
        //public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        //{
        //    List<UpdateDataForClassI> formData = new List<UpdateDataForClassI>();

        //    Dictionary<string, object> resultData = new Dictionary<string, object>();
        //    List<Sheet> sheets = new List<Sheet>
        //    {
        //        SheetHelper.GetSheetById(smartsheet, sheetId1),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId2),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId3),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId4),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId5),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId6),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId7),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId8),
        //        //SheetHelper.GetSheetById(smartsheet, sheetId9)
        //    };
        //    foreach (var sheet in sheets)
        //    {
        //        List<Dictionary<string, object>> rowsData = new List<Dictionary<string, object>>();

        //        foreach (var row in sheet.Rows)
        //        {
        //            if (row.Cells.Any(c => c.DisplayValue == eventId))
        //            {
        //                Dictionary<string, object> rowData = new Dictionary<string, object>();

        //                List<string> columnNames = new List<string>();
        //                foreach (Column column in sheet.Columns)
        //                {
        //                    columnNames.Add(column.Title);
        //                }

        //                for (int i = 0; i < columnNames.Count; i++)
        //                {
        //                    rowData[columnNames[i]] = row.Cells[i].Value;
        //                }

        //                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, row.Id.Value, null);
        //                List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
        //                foreach (var attachment in attachments.Data)
        //                {
        //                    var AID = (long)attachment.Id;
        //                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
        //                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
        //                    {
        //                        { "Name", file.Name },
        //                        { "Url", file.Url }
        //                    };
        //                    attachmentsList.Add(attachmentInfo);
        //                }

        //                rowData["Attachments"] = attachmentsList;
        //                rowsData.Add(rowData);
        //            }
        //        }

        //        if (rowsData.Count > 0)
        //        {
        //            resultData[sheet.Name] = rowsData;
        //        }
        //    }
        //    //foreach(var formdata in resultData)
        //    //{
        //    //    if (formdata.ContainsKey("Event Topic"))
        //    //    {
        //    //        specificData["Event Topic"] = resultData["Event Topic"];
        //    //    }
        //    //}




        //    // string jsonData = JsonConvert.SerializeObject(resultData);

        //    //return Ok(jsonData);
        //    //return Ok(specificData);





        //    return Ok(resultData);
        //    //return Ok(extractedData);




        //}

        [HttpGet("GetDataFromAllSheetsUsingEventIdInPreEvent")]
        public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        {
            Dictionary<string, object> resultData = new();

            List<UpdateDataForClassI> formData = new();

            List<Dictionary<string, object>> eventDetails = new();
            List<Dictionary<string, object>> BrandseventDetails = new();
            List<Dictionary<string, object>> InviteeseventDetails = new();
            List<Dictionary<string, object>> PaneleventDetails = new();
            List<Dictionary<string, object>> SlideKiteventDetails = new();
            List<Dictionary<string, object>> ExpenseeventDetails = new();
            List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> DeviationsattachmentsList = new List<Dictionary<string, object>>();

            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            List<string> columnNames = new List<string>();
            foreach (Column column in sheet1.Columns)
            {
                columnNames.Add(column.Title);
            }

            foreach (var row in sheet1.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> columnsToInclude = new List<string> { "EventDate", "Event Topic", "StartTime", "EndTime", "State", "City" };

                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int i = 0; i < columnNames.Count; i++)
                    {
                        if (columnsToInclude.Contains(columnNames[i]))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, row.Id.Value, null);

                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>();
                    foreach (var attachment in attachments.Data)
                    {
                        var AID = (long)attachment.Id;
                        var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);


                        attachmentInfo[file.Name] = file.Url;




                    }
                    attachmentsList.Add(attachmentInfo);

                    eventDetails.Add(rowData);
                }
            }

            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            List<string> BrandsColumnNames = new List<string>();
            foreach (Column column in sheet2.Columns)
            {
                BrandsColumnNames.Add(column.Title);
            }

            foreach (var row in sheet2.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> BrandscolumnsToInclude = new List<string> { "BrandRequestID", "Brands", "% Allocation", "Project ID" };

                    Dictionary<string, object> BrandsrowData = new Dictionary<string, object>();
                    for (int i = 0; i < BrandsColumnNames.Count; i++)
                    {
                        if (BrandscolumnsToInclude.Contains(BrandsColumnNames[i]))
                        {
                            BrandsrowData[BrandsColumnNames[i]] = row.Cells[i].Value;
                        }
                    }



                    BrandseventDetails.Add(BrandsrowData);
                }
            }
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            List<string> InviteesColumnNames = new List<string>();
            foreach (Column column in sheet3.Columns)
            {
                InviteesColumnNames.Add(column.Title);
            }

            foreach (var row in sheet3.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> InviteescolumnsToInclude = new List<string> { "INV", "Invitee Source", "HCPName", "MISCode", "Employee Code", "LocalConveyance", "BTC/BTE", "LocalConveyance", "Speciality", "Lc Amount Excluding Tax", "LcAmount" };

                    Dictionary<string, object> InviteesrowData = new Dictionary<string, object>();
                    for (int i = 0; i < InviteesColumnNames.Count; i++)
                    {
                        if (InviteescolumnsToInclude.Contains(InviteesColumnNames[i]))
                        {
                            InviteesrowData[InviteesColumnNames[i]] = row.Cells[i].Value;
                        }
                    }



                    InviteeseventDetails.Add(InviteesrowData);
                }
            }

            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            List<string> PanelColumnNames = new List<string>();
            foreach (Column column in sheet4.Columns)
            {
                PanelColumnNames.Add(column.Title);
            }

            foreach (var row in sheet4.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> PanelcolumnsToInclude = new List<string> { "Panelist ID", "SpeakerCode", "TrainerCode", "Tier", "Qualification", "Speciality", "Country",
                        "Rationale", "Speciality", "FCPA Date", "LcAmount", "PresentationDuration", "PanelSessionPreparationDuration", "PanelDiscussionDuration", "QASessionDuration",
                        "BriefingSession", "TotalSessionHours", "HcpRole", "HCPName", "MISCode", "HCP Type", "ExpenseType", "HonorariumRequired", "HonorariumAmount", "Honorarium Amount Excluding Tax",
                        "Travel BTC/BTE", "Mode of Travel", "Travel Excluding Tax", "Travel", "Accomodation Excluding Tax", "Accomodation","Accomodation BTC/BTE",
                        "Local Conveyance Excluding Tax", "LocalConveyance", "LC BTC/BTE",
                        "PAN card name", "Bank Account Number", "IFSC Code", "Bank Name", "Currency", "Other Currency", "Beneficiary Name", "Pan Number", "Global FMV", "Swift Code" };

                    Dictionary<string, object> PanelrowData = new Dictionary<string, object>();
                    for (int i = 0; i < PanelColumnNames.Count; i++)
                    {
                        if (PanelcolumnsToInclude.Contains(PanelColumnNames[i]))
                        {
                            PanelrowData[PanelColumnNames[i]] = row.Cells[i].Value;
                        }
                    }



                    PaneleventDetails.Add(PanelrowData);
                }
            }
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            List<string> SlideKitColumnNames = new List<string>();
            foreach (Column column in sheet5.Columns)
            {
                SlideKitColumnNames.Add(column.Title);
            }

            foreach (var row in sheet5.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> SlideKitcolumnsToInclude = new List<string> { "SlideKit ID", "HCP Name", "MIS", "HcpType", "Slide Kit Type", "SlideKit Document" };

                    Dictionary<string, object> SlideKitrowData = new Dictionary<string, object>();
                    for (int i = 0; i < SlideKitColumnNames.Count; i++)
                    {
                        if (SlideKitcolumnsToInclude.Contains(SlideKitColumnNames[i]))
                        {
                            SlideKitrowData[SlideKitColumnNames[i]] = row.Cells[i].Value;
                        }
                    }



                    SlideKiteventDetails.Add(SlideKitrowData);
                }
            }
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            List<string> ExpenseColumnNames = new List<string>();
            foreach (Column column in sheet6.Columns)
            {
                ExpenseColumnNames.Add(column.Title);
            }

            foreach (var row in sheet6.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> ExpensecolumnsToInclude = new List<string> { "Expenses ID", "Expense", "BTC/BTE", "Amount Excluding Tax", "Amount" };

                    Dictionary<string, object> ExpenserowData = new Dictionary<string, object>();
                    for (int i = 0; i < ExpenseColumnNames.Count; i++)
                    {
                        if (ExpensecolumnsToInclude.Contains(ExpenseColumnNames[i]))
                        {
                            ExpenserowData[ExpenseColumnNames[i]] = row.Cells[i].Value;
                        }
                    }



                    ExpenseeventDetails.Add(ExpenserowData);
                }
            }

            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            List<string> DeviationscolumnNames = new List<string>();
            foreach (Column column in sheet7.Columns)
            {
                DeviationscolumnNames.Add(column.Title);
            }

            foreach (var row in sheet7.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> DeviationscolumnsToInclude = new List<string> { "Deviation Type" };

                    Dictionary<string, object> DeviationsattachmentInfo = new Dictionary<string, object>();
                    for (int i = 0; i < DeviationscolumnNames.Count; i++)
                    {
                        if (DeviationscolumnsToInclude.Contains(DeviationscolumnNames[i]))
                        {
                            var val = row.Cells[i].Value.ToString();
                            var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet7.Id.Value, row.Id.Value, null);

                            
                            foreach (var attachment in attachments.Data)
                            {
                                var AID = (long)attachment.Id;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet7.Id.Value, AID);


                                DeviationsattachmentInfo[val] = file.Url;

                            }
                        }
                    



                    }
                    DeviationsattachmentsList.Add(DeviationsattachmentInfo);

                }
            }








            resultData["eventDetails"] = eventDetails;
            resultData["Files"] = attachmentsList;
            resultData["Brands"] = BrandseventDetails;
            resultData["Invitees"] = InviteeseventDetails;
            resultData["PanelDetails"] = PaneleventDetails;
            resultData["SlideKitSelection"] = SlideKiteventDetails;
            resultData["ExpenseSelection"] = ExpenseeventDetails;
            resultData["Deviation"] = DeviationsattachmentsList;


            return Ok(resultData);
        }

        [HttpPut("UpdateClassIPreEvent")]
        public IActionResult UpdateClassIPreEvent(UpdateAllObjModels formDataList)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new();
            StringBuilder addedInviteesData = new();
            StringBuilder addedMEnariniInviteesData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedSlideKitData = new();
            StringBuilder addedExpences = new();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedInviteesDataNoforMenarini = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;


            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.AmountIncludingTax} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;

            }

            string Expense = addedExpences.ToString();
            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {
                string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType}";
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
                if (formdata.InviteedFrom == "Menarini Employees")
                {
                    string row = $"{addedInviteesDataNoforMenarini}. {formdata.InviteeName}";
                    addedMEnariniInviteesData.AppendLine(row);
                    addedInviteesDataNoforMenarini++;
                }
                else
                {
                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }

            }
            string Invitees = addedInviteesData.ToString();
            string MenariniInvitees = addedMEnariniInviteesData.ToString();
            foreach (var formdata in formDataList.EventRequestHcpRole)
            {

                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.&Acc.Amt: {formdata.Travel + formdata.Accomdation} ";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;

            }
            string HCP = addedHcpData.ToString();



            return Ok();




        }

    }




}
