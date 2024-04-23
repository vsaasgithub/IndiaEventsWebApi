using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.EventTypeSheets
{
    public class Class3
    {
        public DateTime? EventDate { get; set; }
        public string? EventType { get; set; }
        public string? EventName { get; set; }
        public DateTime? EventEndDate { get; set; }
        public string? MeetingType { get; set; }
        public string? Level { get; set; }
        public string? TypeOfParticipation { get; set; }        
        public string? Objective { get; set; }
        public string? VenueName { get; set; }
        public string? State { get; set; }
        public string? City { get; set; }
        public string? TargetParticipants { get; set; }
        public string? VendorName { get; set; }

        public string? InitiatorName { get; set; }
        public int? AdvanceAmount { get; set; }
        public int? TotalExpenseBTC { get; set; }
        public int? TotalExpenseBTE { get; set; }
        public int? TotalHonorariumAmount { get; set; }
        public int? TotalTravelAccommodationAmount { get; set; }
        public int? TotalAccomodationAmount { get; set; }
        public int? TotalBudget { get; set; }
        public int? TotalLocalConveyance { get; set; }
        public int? TotalTravelAmount { get; set; }
        public int? TotalExpense { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? RBMorBMEmail { get; set; }
        public string? SalesHeadEmail { get; set; }
        public string? FinanceHeadEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? MarketingHeadEmail { get; set; }
        public string? FinanceEmail { get; set; }
        public string? ComplianceEmail { get; set; }
        public string? FinanceAccountsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? MedicalAffairsEmail { get; set; }
        public string? Role { get; set; }

        public string? AllIndiaDoctorsInvitedForTheEvent { get; set; }


        public List<string>? FilesUpload { get; set; }
        public string? IsDeviationUpload { get; set; }
        public int? EventOpen45dayscount { get; set; }
        public List<string>? DeviationFiles { get; set; }
    }
}
