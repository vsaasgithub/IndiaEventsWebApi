using IndiaEventsWebApi.Models.EventTypeSheets;
using IndiaEventsWebApi.Models.RequestSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.Draft
{
    public class DraftData
    {
        public Class1? class1 { get; set; }
        public List<EventRequestBrandsList>? RequestBrandsList { get; set; }
        public List<EventRequestInvitees>? EventRequestInvitees { get; set; }
        public List<EventRequestsHcpRole>? EventRequestHcpRole { get; set; }
        public List<EventRequestHCPSlideKit>? EventRequestHCPSlideKits { get; set; }
        public List<EventRequestExpenseSheet>? EventRequestExpenseSheet { get; set; }
    }

    public class PostDraftData
    {
        public DateTime? EventDate { get; set; }
        public string? InitiatorName { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? Isfile { get; set; }
        public List<string>? UploadFiles { get; set; }
    }
    
}
