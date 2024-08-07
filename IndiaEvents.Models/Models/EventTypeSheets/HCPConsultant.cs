﻿using IndiaEvents.Models.Models.RequestSheets;
using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class HCPConsultant
    {
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public DateTime? EventEndDate { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? VenueName { get; set; }
        public string? SponsorshipSocietyName { get; set; }
        public string? Country { get; set; }
        public string? IsAdvanceRequired { get; set; }
        public string? AdvanceAmount { get; set; }
        public string? EventOpen30days { get; set; }
        public string? EventOpen30dayscount { get; set; }
        public string? EventWithin7days { get; set; }
        public string? FcpaFile { get; set; }
        public string? BrochureFile { get; set; }
        public string? AggregateDeviation { get; set; }
        public string? RBMorBM { get; set; }
        public string? Sales_Head { get; set; }
        public string? FinanceHead { get; set; }
        public string? Marketing_Head { get; set; }
        //public string? MedicalAffairsEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? MarketingCoordinatorEmail { get; set; }
        public string? TotalExpenseBTC { get; set; }
        public string? TotalExpenseBTE { get; set; }

        public string? MedicalAffairsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? ComplianceEmail { get; set; }
        //public string? FinanceTreasuryEmail { get; set; }
        public string? FinanceAccountsEmail { get; set; }
        public string? Role { get; set; }

        public string? Finance { get; set; }
        public string? IsFilesUpload { get; set; }
        public List<string>? Files { get; set; }

        public string? InitiatorName { get; set; }
        //public string? Role { get; set; }
        public string? Initiator_Email { get; set; }
        public List<string>? AggregateDeviationFiles { get; set; }

    }

    public class HCPList
    {
        public string? HcpName { get; set; }
        public string? MisCode { get; set; }
        public string? HcpType { get; set; }
        public string? TravelAmount { get; set; }
        public string? AccomAmount { get; set; }
        public string? LcAmount { get; set; }
        //public string? LcBtcorBte { get; set; }
        //public string? TravelBtcorBte { get; set; }
        //public string? AccomodationBtcorBte { get; set; }
        //public int? HonarariumAmountExcludingTax { get; set; }
        //public int? TravelExcludingTax { get; set; }
        //public int? AccomdationExcludingTax { get; set; }
        //public int? LocalConveyanceExcludingTax { get; set; }

        public string? RegistrationAmount { get; set; }
        public string? BudgetAmount { get; set; }
        public string? Legitimate { get; set; }
        public string? Objective { get; set; }
        public string? Rationale { get; set; }
        public DateTime? Fcpadate { get; set; }
        public string? ExpenseType { get; set; }
        public string? IsUpload { get; set; }
        public List<string>? FilesToUpload { get; set; }



    }
    public class HCPListForHcpConsuktant
    {
        public string? HcpName { get; set; }
        public string? MisCode { get; set; }
        public string? HcpType { get; set; }
        public double TrainTravelAmountIncludingTax { get; set; }
        public string? TrainTravelBtcBte { get; set; }
        public double TrainTravelAmountExcludingTax { get; set; }
        public double AirTravelAmountIncludingTax { get; set; }
        public double AirTravelAmountExcludingTax { get; set; }
        public string? AirTravelBtcBte { get; set; }
        public double RoadTravelAmountIncludingTax { get; set; }
        public double RoadTravelAmountExcludingTax { get; set; }
        public string? RoadTravelBtcBte { get; set; }
        public double AccomAmountIncludingTax { get; set; }
        public double AccomAmountExcludingTax { get; set; }
        public string? AccomodationBtcorBte { get; set; }
        public double LcAmountIncludingTax { get; set; }
        public double LcAmountExcludingTax { get; set; }
        public string? LcBtcorBte { get; set; }
        public double RegistrationAmountIncludingTax { get; set; }
        public double RegistrationAmountExcludingTax { get; set; }
        public string? RegistrationAmountBtcBte { get; set; }
        public double BudgetAmount { get; set; }
        public string Legitimate { get; set; }
        public string Objective { get; set; }
        public string Rationale { get; set; }
        public DateTime? Fcpadate { get; set; }
        public string? ExpenseType { get; set; }
        public string? IsUpload { get; set; }
        public List<string>? FilesToUpload { get; set; }



    }
    public class ExpenseList
    {

        public string? Expense { get; set; }
        public string? BTC_BTE { get; set; }

        public string? RegstAmount { get; set; }
        public double? RegstAmountExcludingTax { get; set; }
        public string? ExpenseAmount { get; set; }
        public double? ExpenseAmountExcludingTax { get; set; }
        public string? BtcAmount { get; set; }
        public string? BteAmount { get; set; }




    }

    public class ExpenseListForHcpConsultant
    {

        public string? Expense { get; set; }
        public string? BTC_BTE { get; set; }
        public string? MisCode { get; set; }

        public string? RegstAmount { get; set; }
        public double? RegstAmountExcludingTax { get; set; }
        public string? ExpenseAmount { get; set; }
        public double? ExpenseAmountExcludingTax { get; set; }
        public string? BtcAmount { get; set; }
        public string? BteAmount { get; set; }
        public string? BudgetAmount { get; set; }




    }
    public class HCPConsultantPayload
    {
        public HCPConsultant HcpConsultant { get; set; }
        public List<EventRequestBrandsList>? BrandsList { get; set; }
        public List<ExpenseListForHcpConsultant>? ExpenseSheet { get; set; }
        public List<HCPListForHcpConsuktant>? HcpList { get; set; }
        public string? IsDeviationUpload { get; set; }
        public List<EventRequestDeviationsData>? DeviationDetails { get; set; }
    }

    public class HCPfollow_upsheet
    {
        public string? HCPName { get; set; }
        public string? MisCode { get; set; }
        public string? GO_NGO { get; set; }
        public string? Country { get; set; }
        public string? How_many_days_since_the_parent_event_completes { get; set; }
        public string? Follow_up_Event { get; set; }
        public DateTime? Follow_up_Event_Date { get; set; }
        public DateTime? Event_Date { get; set; }
        public string? AgreementFile { get; set; }
    }
}
