﻿namespace IndiaEventsWebApi.Models.RequestSheets
{
    public class EventRequestExpenseSheet
    {
        public string? EventId { get; set; }
        public string? Expense { get; set; }
        public string? Amount { get; set; }
        public string? AmountExcludingTax { get; set; }
        public double? ExcludingTaxAmount { get; set; }
        public string? BtcorBte { get; set; }
        public string? BtcAmount { get; set; }
        public string? BteAmount { get; set; }
        public string? BudgetAmount { get; set; }
    }

    public class AddNewExpense
    {
        public string? EventId { get; set; }
        public string? ExpenseType { get; set; }
        //public string? Amount { get; set; }
        public double? AmountIncludingTax { get; set; }
        public double? ExcludingTaxAmount { get; set; }
        public string? BtcorBte { get; set; }
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public string? VenueName { get; set; }
        public string? IsExpenseAddedInPostSettlement { get; set; }
        public DateTime? EventStartDate { get; set; }
        public DateTime? EventEndDate { get; set; }
    }
}
