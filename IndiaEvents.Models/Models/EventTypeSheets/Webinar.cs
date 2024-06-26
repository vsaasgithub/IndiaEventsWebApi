﻿using IndiaEvents.Models.Models.RequestSheets;
using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class Webinar
    {
        public string? EventId { get; set; }
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public DateTime? EventEndDate { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? Meeting_Type { get; set; }
        public string? BTEExpenseDetails { get; set; }
        //public string? VenueName { get; set; }
        //public string? City { get; set; }
        //public string? State { get; set; }
        public string? BrandName { get; set; }
        public string? PercentAllocation { get; set; }
        public string? ProjectId { get; set; }
        public string? HCPRole { get; set; }
        public string? IsAdvanceRequired { get; set; }
        public string? AdvanceAmount { get; set; }
        public string? TotalExpenseBTC { get; set; }
        public string? TotalExpenseBTE { get; set; }
        public string? EventOpen30days { get; set; }
        public string? EventOpen30dayscount { get; set; }
        public string? EventWithin7days { get; set; }
        public string? FB_Expense_Excluding_Tax { get; set; }
        public string? RBMorBM { get; set; }
        public string? FinanceHead { get; set; }


        //public string? MedicalAffairsEmail { get; set; }
        public string? MedicalAffairsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? ComplianceEmail { get; set; }
        //public string? FinanceTreasuryEmail { get; set; }
        public string? FinanceAccountsEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? MarketingCoordinatorEmail { get; set; }
        public string? Role { get; set; }


        public string? Sales_Head { get; set; }
        public string? Marketing_Head { get; set; }
        public string? Finance { get; set; }
        public string? InitiatorName { get; set; }
        public string? Initiator_Email { get; set; }
        public List<string>? Files { get; set; }
        public string? IsDeviationUpload { get; set; }
        public List<EventRequestDeviationsData>? DeviationDetails { get; set; }
        //public List<string>? DeviationFiles { get; set; }
        //public string? Role { get; set; }
    }

    public class WebinarPayload
    {
        public Webinar? Webinar { get; set; }
        public List<EventRequestBrandsList>? RequestBrandsList { get; set; }
        public List<EventRequestInvitees>? EventRequestInvitees { get; set; }
        public List<EventRequestsHcpRole>? EventRequestHcpRole { get; set; }
        public List<EventRequestHCPSlideKit>? EventRequestHCPSlideKits { get; set; }
        public List<EventRequestExpenseSheet>? EventRequestExpenseSheet { get; set; }
        // public IFormFile? formFile { get; set; }
    }
  
}
