using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.FMV
{
    [Route("api/[controller]")]
    [ApiController]
    public class FMVController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public FMVController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }

        [HttpGet("GetfmvColumnValue")]
        public IActionResult GetfmvColumnValue(string specialty, string columnTitle)
        {
            int defaultval = 0;
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:fmv").Value;
            long.TryParse(sheetId, out long parsedSheetId);
            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "Speciality", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, columnTitle, StringComparison.OrdinalIgnoreCase));
            if (targetColumn != null && SpecialityColumn != null)
            {
                Row targetRow = sheet.Rows.FirstOrDefault(row => row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == specialty));

                if (targetRow != null)
                {
                    var columnValue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value;
                    if (columnValue != null)
                    {
                        return Ok(columnValue);
                    }

                    else
                    {
                        //var notFoundObject = new
                        //{
                        //    Message = $"Value not found for {specialty} in {columnTitle} column.",
                        //    code = 404
                        //};

                        //return NotFound(notFoundObject);
                        return Ok(defaultval);
                    }
                }

                else
                {
                    //var notFoundObject = new
                    //{
                    //    Message = $"Speciality '{specialty}' not found.",
                    //    code = 404
                    //};

                    //return NotFound(notFoundObject);
                    return Ok(defaultval);
                }
            }

            else
            {
                //var notFoundObject = new
                //{
                //    Message = $"Column '{columnTitle}' not found.",
                //    code = 404
                //};

                //return NotFound(notFoundObject);
                return Ok(defaultval);
            }
        }


        [HttpGet("GetFMVData")]
        public IActionResult GetFMVData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:fmv").Value;
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
    }
}
