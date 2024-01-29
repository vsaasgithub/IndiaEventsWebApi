﻿using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class HonorariumPayment
    {
        
        public string EventId { get; set; }
        public string EventType { get; set; }
        public string EventTopic { get; set; }
        public string HonarariumSubmitted { get; set; }
        public DateTime EventDate { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string VenueName { get; set; }
        public string TotalTravelAndAccomodationSpend { get; set; }
        public string TotalHonorariumSpend { get; set; }
        public string TotalSpend { get; set; }
        public string TotalLocalConveyance { get; set; }
        public string Brands { get; set; }
        public string Invitees { get; set; }
        public string Panelists { get; set; }
        public string InitiatorName { get; set; }
        public string InitiatorEmail { get; set; }
        public string RBMorBM { get; set; }
        public string Compliance { get; set; }
        public string FinanceAccounts { get; set; }
        public string FinanceTreasury { get; set; }
        public string slideKits { get; set; }
        public string Expenses { get; set; }
        public string TotalTravelSpend { get; set; }
        public string TotalAccomodationSpend { get; set; }
        public string TotalExpenses { get; set; }
    }


    public class HCPDetails
    {
        public string HcpName { get; set; }
        public string HcpRole { get; set; }
        public string MisCode { get; set; }
        public string GOorNGO { get; set; }
        public string IsInclidingGst { get; set; }
        public string AgreementAmount { get; set; }
    }
    //public class Branddetails
    //{
    //    public string BrandName { get; set; }
    //    public string PercentAllocation { get; set; }
    //    public string ProjectId { get; set; }
    //}
    //public class Invitees
    //{
       
    //    public string InviteeName { get; set; }
    //    public string MISCode { get; set; }
    //    public string LocalConveyance { get; set; }
        
    //}
    //public  class Panalist
    //{
    //    public string HcpRole { get; set; }
    //    public string HcpName { get; set; }
    //    public string HonarariumAmount { get; set; }
    //    public string TravelAmount { get; set; }
    //    public string AccomdationAmount { get; set; }

    //}
    //public class ExpenseSheet
    //{
        
    //    public string Expense { get; set; }
    //    public string Amount { get; set; }
    //    public string AmountExcludingTax { get; set; }
    //    public string BtcorBte { get; set; }
      
    //}

    //public class HCPSlideKit
    //{
        
    //    public string MIS { get; set; }
    //    public string SlideKitType { get; set; }
    //    public string SlideKitDocument { get; set; }


    //}

    public class HonorariumPaymentList
    {
        public HonorariumPayment RequestHonorariumList { get; set; }
        public List<HCPDetails>? HcpRoles { get; set; }
        //public List<Branddetails>? BrandDetails { get; set; }
        //public List<Invitees>? Invitees { get; set; }
        //public List<Panalist>? panalist { get; set; }
        //public List<ExpenseSheet>? expenseSheet { get; set; }
        //public List<HCPSlideKit>? hCPSlideKits { get; set; }


    }
}

