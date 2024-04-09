﻿namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class EventSettlement
    {
        public List<Invitee>? Invitee { get; set; }
        public List<ExpenseSheet>? expenseSheets { get; set; }

        public string? EventId { get; set; }
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public string? StartTime { get; set; }
        public string? InitiatorName { get; set; }
        public string? EndTime { get; set; }
        public string? VenueName { get; set; }
        public string? State { get; set; }
        public string? City { get; set; }
        public string? Attended { get; set; }
        public string? Brands { get; set; }
        public string? Panalists { get; set; }
        public string? SlideKits { get; set; }
        public string? TotalExpense { get; set; }
        public string? Advance { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? RBMorBM { get; set; }
        public string? SalesHead { get; set; }
        public string? MarkeringHead { get; set; }
        public string? Compliance { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? FinanceTreasuryEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? Role { get; set; }

        public string? FinanceAccounts { get; set; }
        public string? FinanceTreasury { get; set; }
        public string? MedicalAffairsHead { get; set; }
        public string? FinanceHead { get; set; }
        public string? EventOpen30Days { get; set; }
        public string? EventLessThan5Days { get; set; }
        public string? PostEventSubmitted { get; set; }
        public string? IsAdvanceRequired { get; set; }
        public string? totalInvitees { get; set; }
        public string? TotalAttendees { get; set; }
        public string? TotalTravelSpend { get; set; }
        public string? TotalAccomodationSpend { get; set; }
        public string? TotalExpenses { get; set; }
        public string? TotalTravelAndAccomodationSpend { get; set; }
        public string? TotalHonorariumSpend { get; set; }
        public string? TotalSpend { get; set; }
        public string? TotalActuals { get; set; }
        public string? AdvanceUtilizedForEvents { get; set; }
        public string? PayBackAmountToCompany { get; set; }
        public string? AdditionalAmountNeededToPayForInitiator { get; set; }
        public string? TotalLocalConveyance { get; set; }
        public string? IsDeviationUpload { get; set; }


        public List<string>? Files { get; set; }
        public List<string>? DeviationFiles { get; set; }



    }

    public class Invitee
    {

        public string? InviteeName { get; set; }
        public string? MISCode { get; set; }
        public string? LocalConveyance { get; set; }
        public string? BtcorBte { get; set; }
        public string? LcAmount { get; set; }
    }
    public class ExpenseSheet
    {

        public string? Expense { get; set; }
        public string? Amount { get; set; }
        public string? AmountExcludingTax { get; set; }
        public string? BtcorBte { get; set; }

    }

    public class HandsOnPost
    {

        public string? EventId { get; set; }
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public string? StartTime { get; set; }
        public string? InitiatorName { get; set; }
        public string? EndTime { get; set; }
        public string? VenueName { get; set; }
        public string? State { get; set; }
        public string? City { get; set; }
        public string? Attended { get; set; }
        public string? Brands { get; set; }
        public string? Panalists { get; set; }
        public string? SlideKits { get; set; }
        public string? TotalExpense { get; set; }
        public string? Advance { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? RBMorBM { get; set; }
        public string? SalesHead { get; set; }
        public string? MarkeringHead { get; set; }
        public string? Compliance { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? FinanceTreasuryEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? Role { get; set; }
        public string? FinanceAccounts { get; set; }
        public string? FinanceTreasury { get; set; }
        public string? MedicalAffairsHead { get; set; }
        public string? FinanceHead { get; set; }
        public string? EventOpen30Days { get; set; }
        public string? EventLessThan5Days { get; set; }
        public string? PostEventSubmitted { get; set; }
        public string? IsAdvanceRequired { get; set; }
        public string? TotalInvitees { get; set; }
        public string? TotalAttendees { get; set; }
        public string? TotalTravelSpend { get; set; }
        public string? TotalAccomodationSpend { get; set; }
        public string? TotalExpenses { get; set; }
        public string? TotalTravelAndAccomodationSpend { get; set; }
        public string? TotalHonorariumSpend { get; set; }
        public string? TotalSpend { get; set; }
        public string? TotalActuals { get; set; }
        public string? AdvanceUtilizedForEvents { get; set; }
        public string? PayBackAmountToCompany { get; set; }
        public string? AdditionalAmountNeededToPayForInitiator { get; set; }
        public string? TotalLocalConveyance { get; set; }
        public string? IsDeviationUpload { get; set; }


        public List<string>? Files { get; set; }
        public List<string>? DeviationFiles { get; set; }
        public List<UpdatePanelDetails>? PanelData { get; set; }
        public List<UpdateInviteeDetails>? InviteesData { get; set; }
        public List<UpdateExpenseDetails>? ExpenseData { get; set; }
        public List<UpdateSlideKitDetails>? SlideKitData { get; set; }

    }


    public class UpdatePanelDetails
    {
        public string? PanelId { get; set; }
        public string? IsAttended { get; set; }
        public int? ActualAccomodationAmount { get; set; }
        public int? ActualTravelAmount { get; set; }
        public int? ActualLCAmount { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }
    public class UpdateInviteeDetails
    {
        public string? InviteeId { get; set; }
        public string? IsAttended { get; set; }
        public int? ActualAmount { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }
    public class UpdateExpenseDetails
    {
        public string? ExpenseId { get; set; }
        public int? ActualAmount { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }
    public class UpdateSlideKitDetails
    {
        public string? SlideKitId { get; set; }
        public string? TrainerOrHcpName { get; set; }
        public string? ProductName { get; set; }
        public string? IndicationsDone { get; set; }
        public string? BatchNumber { get; set; }
        public string? SubjectNameandSurName { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }

    public class IssuedQuantityDetails
    {
        public string? Id { get; set; }
        public string? ProductName { get; set; }
        public string? BatchNumber { get; set; }
        public int? IssuedQuantity { get; set; }
    }

    public class SamplesConsumedDetails
    {
        public string? Id { get; set; }
        public int? TrainerName { get; set; }
        public string? ProductName { get; set; }
        public string? SamplesUsed { get; set; }
    }

}

