using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
   
    public class FCPAMasterController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public FCPAMasterController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetHCPDataUsingNameAndMISCode")]
        public IActionResult GetHCPDataUsingNameAndMISCode(string misCode)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string[] sheetIds = {
            configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
            configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
            configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
            configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
            configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
        };
            foreach (string i in sheetIds)
            {
                long.TryParse(i, out long p);
                Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);

                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MISCode");

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

                        return Ok("True");
                    }

                }
            }
            return Ok("False");
        }

    }
}

