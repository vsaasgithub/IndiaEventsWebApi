﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.EventTypeSheets
{
    public class UpdateAttendees
    {
        public string EventId { get; set; }
        public List<UpdatePanelDetails> PanelData { get; set; }
        public List<UpdateInviteeDetails> InviteesData { get; set; }
        public List<UpdateExpenseDetails> ExpenseData { get; set; }
        public List<UpdateSlideKitDetails> SlideKitData { get; set; }
    }
    public class UpdatePanelDetails
    {
        public string? PanelId { get; set; }
        public string? IsAttended { get; set; }
        public double? ActualAccomodationAmount { get; set; }
        public double? ActualTravelAmount { get; set; }
        public double? ActualLCAmount { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }
    public class UpdateInviteeDetails
    {
        public string? InviteeId { get; set; }
        public string? IsAttended { get; set; }
        public double? ActualAmount { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }
    public class UpdateExpenseDetails
    {
        public string? ExpenseId { get; set; }
        public double? ActualAmount { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }
    public class UpdateSlideKitDetails
    {
        public string? SlideKitId { get; set; }
        public string? ProductName { get; set; }
        public string? IndicationsDone { get; set; }
        public string? BatchNumber { get; set; }
        public string? SubjectNameandSurName { get; set; }
        public string? IsUploadDocument { get; set; }
        public List<string>? UploadDocument { get; set; }
    }


}
