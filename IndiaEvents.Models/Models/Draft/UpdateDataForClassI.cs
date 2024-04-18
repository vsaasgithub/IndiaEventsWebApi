using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.Draft
{
    public class UpdateDataForClassI
    {
       
        public EventDetails EventDetails { get; set; }
        public List<BrandSelection> BrandSelection { get; set; }
        public List<PanelSelection> PanelSelection { get; set; }
        public List<SlideKitSelection> SlideKitSelection { get; set; }
        public List<InviteeSelection> InviteeSelection { get; set; }
        public List<ExpenseSelection> ExpenseSelection { get; set; }
        public Files files { get; set; }
    }

    public class BrandSelection
    {
        public string brandName { get; set; }
        public int percentageAllocation { get; set; }
        public string projectId { get; set; }
    }

    public class EventDetails
    {
        public string EventTopic { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string VenueName { get; set; }
        public string State { get; set; }
        public string City { get; set; }
    }

    public class ExpenseSelection
    {
        public string ExpenseName { get; set; }
        public string ExpenseType { get; set; }
        public int ExpenseAmountIncludingTax { get; set; }
        public int ExpenseAmountExcludingTax { get; set; }
        public string FoodAndBeveragedDeviationUrl { get; set; }
    }

    public class Files
    {
        public string AgendaFileUrl { get; set; }
        public string InvitationFileUrl { get; set; }
    }

    public class InviteeSelection
    {
        public string InviteeFrom { get; set; }
        public string Name { get; set; }
        public string MisCode { get; set; }
        public string EmployeeCode { get; set; }
        public string isLocalConveyance { get; set; }
        public string localConveyanceType { get; set; }
        public string speciality { get; set; }
        public int localConveyanceAmountIncludingTax { get; set; }
        public int localConveyanceAmountExcludingTax { get; set; }
    }

    public class PanelSelection
    {
        public string SpeakerCode { get; set; }
        public string TrainerCode { get; set; }
        public string Speciality { get; set; }
        public string Tier { get; set; }
        public string Qualification { get; set; }
        public string Country { get; set; }
        public string Rationale { get; set; }
        public string nocFileUrl { get; set; }
        public string fcpaIssueDate { get; set; }
        public string fcpaFile { get; set; }
        public int PresentationDuration { get; set; }
        public int PanelSessionPreperationDuration { get; set; }
        public int PanelDisscussionDuration { get; set; }
        public int QaSessionDuration { get; set; }
        public int BriefingSession { get; set; }
        public int TotalSessionHours { get; set; }
        public string HcpRole { get; set; }
        public string HcpName { get; set; }
        public string MisCode { get; set; }
        public string GOorNGO { get; set; }
        public string ExpenseType { get; set; }
        public string HonorariumRequired { get; set; }
        public int HonarariumAmountIncludingTax { get; set; }
        public int HonarariumAmountExcludingTax { get; set; }
        public int TravelAmountIncludingTax { get; set; }
        public int TravelExcludingTax { get; set; }
        public string TravelBtcorBte { get; set; }
        public int LocalConveyanceIncludingTax { get; set; }
        public int LocalConveyanceExcludingTax { get; set; }
        public string LcBtcorBte { get; set; }
        public int AccomdationIncludingTax { get; set; }
        public int AccomdationExcludingTax { get; set; }
        public string AccomodationBtcorBte { get; set; }
        public string aggregateDeviationFileUrl { get; set; }
        public string honorariumDeviationFileUrl { get; set; }
        public string PanCardName { get; set; }
        public string BankAccountNumber { get; set; }
        public string IFSCCode { get; set; }
        public string BankName { get; set; }
        public string Currency { get; set; }
        public string OtherCurrencyType { get; set; }
        public string BeneficiaryName { get; set; }
        public string PanNumber { get; set; }
        public string isGlobalFMVCheck { get; set; }
        public string SwiftCode { get; set; }
        public string PanCardFileUrl { get; set; }
        public string ChequeFileUrl { get; set; }
        public string TaxResidentialFileUrl { get; set; }
    }

    public class PreEventCheck
    {
        [JsonProperty("45DaysDeviationFileUrl")]
        public string _45DaysDeviationFileUrl { get; set; }
        public string eventDate { get; set; }

        [JsonProperty("5DaysDeviationFileUrl")]
        public string _5DaysDeviationFileUrl { get; set; }
    }

    public class Root
    {
        public PreEventCheck preEventCheck { get; set; }
        public EventDetails eventDetails { get; set; }
        public List<BrandSelection> brandSelection { get; set; }
        public List<PanelSelection> panelSelection { get; set; }
        public List<SlideKitSelection> slideKitSelection { get; set; }
        public List<InviteeSelection> inviteeSelection { get; set; }
        public List<ExpenseSelection> expenseSelection { get; set; }
        public Files files { get; set; }
    }

    public class SlideKitSelection
    {
        public string hcpName { get; set; }
        public string misCode { get; set; }
        public string hcpType { get; set; }
        public string slideKitType { get; set; }
        public string slideKitOption { get; set; }
        public string slideKitFileUrl { get; set; }
    }


}
