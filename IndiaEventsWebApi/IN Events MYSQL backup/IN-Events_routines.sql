CREATE DATABASE  IF NOT EXISTS `IN-Events` /*!40100 DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci */ /*!80016 DEFAULT ENCRYPTION='N' */;
USE `IN-Events`;
-- MySQL dump 10.13  Distrib 8.0.36, for Win64 (x86_64)
--
-- Host: 10.9.128.92    Database: IN-Events
-- ------------------------------------------------------
-- Server version	8.0.25

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!50503 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Dumping events for database 'IN-Events'
--

--
-- Dumping routines for database 'IN-Events'
--
/*!50003 DROP PROCEDURE IF EXISTS `Class1Preevent` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `Class1Preevent`(ApproverPreEventURL varchar(500),FinanceTreasuryURL varchar(500),InitiatorURL varchar(500),
EventTopic text,EventType varchar(100),EventDate date,StartTime varchar(15),EndTime varchar(15),MeetingType text,Brands text,Expenses text,
Panelists text,Invitees text,MIPLInvitees text,SlideKits text,IsAdvanceRequired varchar(5),EventOpen30days varchar(5),EventWithin7days varchar(5),
InitiatorName varchar(100),AdvanceAmount Decimal(15,4),TotalExpenseBTC Decimal(15,4),TotalExpenseBTE Decimal(15,4),TotalHonorariumAmount Decimal(15,4),
TotalTravelAmount Decimal(15,4),TotalTravelAccommodationAmount Decimal(15,4),TotalAccommodationAmount Decimal(15,4),BudgetAmount Decimal(15,4),TotalLocalConveyance Decimal(15,4),
TotalExpense Decimal(15,4),InitiatorEmail varchar(200),RBMBM varchar(200),SalesHead varchar(200),SalesCoordinator varchar(200),MarketingCoordinator varchar(200),
MarketingHead varchar(200),Compliance varchar(200),FinanceAccounts varchar(200),FinanceTreasury varchar(200),ReportingManager varchar(200),
1UpManager varchar(200),MedicalAffairsHead varchar(200),BTEExpenseDetails text,AttachmentPaths text,webinarRole varchar(50),
VenueName text ,City varchar(100), State varchar(100)
)
BEGIN

Insert into EventRequestsWeb
(`Approver Pre Event URL`,`Finance Treasury URL`,`Initiator URL`,`Event Topic`,`Event Type`,`Event Date`,`Start Time`,`End Time`,
`Meeting Type`,`Brands`,`Expenses`,`Panelists`,`Invitees`,`MIPL Invitees`,`SlideKits`,`IsAdvanceRequired`,`EventOpen30days`,`EventWithin7days`,`Initiator Name`,
`Advance Amount`,`Total Expense BTC`,`Total Expense BTE`,`Total Honorarium Amount`,`Total Travel Amount`,`Total Travel & Accommodation Amount`,`Total Accommodation Amount`,
`Budget Amount`,`Total Local Conveyance`,`Total Expense`,`Initiator Email`,`RBM/BM`,`Sales Head`,`Sales Coordinator`,`Marketing Coordinator`,
`Marketing Head`,`Compliance`,`Finance Accounts`,`Finance Treasury`,`Reporting Manager`,`1 Up Manager`,`Medical Affairs Head`,`BTE Expense Details`,`AttachmentPaths`,`Created On`,
`syncdone`,`Role`,`Venue Name`,`City`,`State`) 
values
(ApproverPreEventURL,FinanceTreasuryURL,InitiatorURL,EventTopic,EventType,EventDate,StartTime,EndTime,MeetingType,
Brands,Expenses,Panelists,Invitees,MIPLInvitees,SlideKits,IsAdvanceRequired,EventOpen30days,EventWithin7days,InitiatorName,AdvanceAmount,TotalExpenseBTC,
TotalExpenseBTE,TotalHonorariumAmount,TotalTravelAmount,TotalTravelAccommodationAmount,TotalAccommodationAmount,BudgetAmount,TotalLocalConveyance,TotalExpense,
InitiatorEmail,RBMBM,SalesHead,SalesCoordinator,MarketingCoordinator,MarketingHead,Compliance,FinanceAccounts,FinanceTreasury,ReportingManager,
1UpManager,MedicalAffairsHead,BTEExpenseDetails,AttachmentPaths,now(),0,webinarRole,VenueName,City,State);
SELECT LAST_INSERT_ID() as ID;
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `GetAllEventSyncData` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `GetAllEventSyncData`()
BEGIN
CREATE TEMPORARY TABLE IF NOT EXISTS TempEventRequest AS
(select ID,`Approver Pre Event URL`,`Finance Treasury URL`,`Initiator URL`,`Event Topic`,`Event Type`,`Event Date`,`Start Time`,`End Time`,
`Meeting Type`,`Brands`,`Expenses`,`Panelists`,`Invitees`,`MIPL Invitees`,`SlideKits`,`IsAdvanceRequired`,`EventOpen30days`,`EventWithin7days`,`Initiator Name`,
`Advance Amount`,`Total Expense BTC`,`Total Expense BTE`,`Total Honorarium Amount`,`Total Travel Amount`,`Total Travel & Accommodation Amount`,`Total Accommodation Amount`,
`Budget Amount`,`Total Local Conveyance`,`Total Expense`,`Initiator Email`,`RBM/BM`,`Sales Head`,`Sales Coordinator`,`Marketing Coordinator`,
`Marketing Head`,`Compliance`,`Finance Accounts`,`Finance Treasury`,`Reporting Manager`,`1 Up Manager`,`Medical Affairs Head`,`BTE Expense Details`,`AttachmentPaths`,`Role`,`Class III Event Code`,
CASE `End Date` 
when '0000-00-00' then ''
when null then ''
else `End Date` end as `End Date`,`Venue Name`,`City`,`State` from EventRequestsWeb where syncdone = 0);

select * from TempEventRequest ;

Select `HcpRole`,`MISCode`,`Travel`,`TotalSpend`,`Accomodation`,`LocalConveyance`,`SpeakerCode`,`TrainerCode`,`HonorariumRequired`,
`AgreementAmount`,`HonorariumAmount`,`Speciality`,`Event Topic`,`Event Type`,`Event Date Start`,`Event End Date`,`HCPName`,`PAN card name`,`ExpenseType`,
`Bank Account Number`,`Bank Name`,`IFSC Code`,CASE `FCPA Date` when '0000-00-00' then '' else  `FCPA Date` end as `FCPA Date` ,`Currency`,`Honorarium Amount Excluding Tax`,`Travel Excluding Tax`,`Accomodation Excluding Tax`,
`Local Conveyance Excluding Tax`,`LC BTC/BTE`,`Travel BTC/BTE`,`Accomodation BTC/BTE`,`Mode of Travel`,`Other Currency`,`Beneficiary Name`,`Pan Number`,
`Other Type`,`Tier`,`HCP Type`,`PresentationDuration`,`PanelSessionPreparationDuration`,`PanelDiscussionDuration`,`QASessionDuration`,
`BriefingSession`,`TotalSessionHours`,`Rationale `,`EventId/EventRequestId`,AttachmentPaths 
from EventRequestPanelDetails where `EventId/EventRequestId` in (Select ID from TempEventRequest);

select `% Allocation`,`Brands`,`Project ID`,`EventId/EventRequestId` 
from EventRequestsBrandsList where `EventId/EventRequestId` in(Select ID from TempEventRequest);

select `HCPName`,`Designation`,`Employee Code`,`LocalConveyance`,`BTC/BTE`,`LcAmount`,`Lc Amount Excluding Tax`,`EventId/EventRequestId`,
`Invitee Source`,`HCP Type`,`Speciality`,`MISCode`,`Event Topic`,`Event Type`,`Event Date Start`,`Event End Date` 
from EventRequestInvitees where `EventId/EventRequestId` in(Select ID from TempEventRequest);

select `MIS`,`Slide Kit Type`,`SlideKit Document`,`EventId/EventRequestId`,AttachmentPaths 
from EventRequestHCPSlideKitDetails where `EventId/EventRequestId` in(Select ID from TempEventRequest);

select `Expense`,`EventId/EventRequestID`,`AmountExcludingTax?`,`Amount Excluding Tax`,`Amount`,
`BTC/BTE`,`BudgetAmount`,`BTCAmount`,`BTEAmount`,`Event Topic`,`Event Type`,`Event Date Start`,`Event End Date` 
from EventRequestExpensesSheet where `EventId/EventRequestId` in(Select ID from TempEventRequest);

select `EventId/EventRequestId`,`Event Topic`,`Event Type`,`Event Date`,`Start Time`,`End Time`,
`MIS Code`,`HCP Name`,`Honorarium Amount`,`Travel & Accommodation Amount`,`Other Expenses`,`Sales Head`,`Finance Head`,`Initiator Name`,
`Initiator Email`,`Sales Coordinator`,`Deviation Type`,`EventOpen45days`,`Outstanding Events`,`EventWithin5days`,`PRE-F&B Expense Excluding Tax`,
`Travel/Accomodation 3,00,000 Exceeded Trigger`,`Trainer Honorarium 12,00,000 Exceeded Trigger`,`HCP Honorarium 6,00,000 Exceeded Trigger`,
CASE `End Date` 
when '0000-00-00' then ''
when null then ''
else  `End Date` end as `End Date`,AttachmentPaths
from Deviation_Process where `EventId/EventRequestId` in(Select ID from TempEventRequest);

Drop temporary table TempEventRequest;
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `GetHCPMaster` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `GetHCPMaster`()
BEGIN
Select ID,LastName,FirstName,HCPName,`HCP Type`,`Employee Group`,`External ID`,`MisCode`,`Company Name`,`Medical Council Registration`,`Speciality`,`FCPA Sign Off Date`,`FCPA Expiry Date`,`FCPA Valid?` from HCPMaster;

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HandsonPreevent` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HandsonPreevent`(EventTopic text,EventType varchar(100),
EventDate date,StartTime varchar(15),EndTime varchar(15),VenueName text ,City varchar(100), State varchar(100),
InitiatorName varchar(100),AdvanceAmount Decimal(15,4),TotalExpenseBTC Decimal(15,4),TotalExpenseBTE Decimal(15,4),TotalHonorariumAmount Decimal(15,4),
TotalTravelAmount Decimal(15,4),TotalTravelAccommodationAmount Decimal(15,4),TotalAccommodationAmount Decimal(15,4),BudgetAmount Decimal(15,4),TotalLocalConveyance Decimal(15,4),
TotalExpense Decimal(15,4),InitiatorEmail varchar(200),RBMBM varchar(200),SalesHead varchar(200),SalesCoordinator varchar(200),MarketingCoordinator varchar(200),
MarketingHead varchar(200),Compliance varchar(200),FinanceAccounts varchar(200),FinanceTreasury varchar(200),ReportingManager varchar(200),
1UpManager varchar(200),MedicalAffairsHead varchar(200),Panelists text,Invitees text,MIPLInvitees text,Brands text,Expenses text,
SlideKits text,BTEExpenseDetails text,AttachmentPaths text,webinarRole varchar(50),ProductBrand text,ModeofTraining text,HOTWebinarVendorName text,
HOTWebinarType text,VenueSelectionChecklist text,EmergencySupport text,EmergencyContactNo text,FacilityCharges Decimal(15,4),FacilityChargesBTCBTE Decimal(15,4),
FacilityChargesExcludingTax Decimal(15,4),TotalFacilityChargesIncludingTax Decimal(15,4),AnesthetistRequiredq varchar(10),
AnesthetistBTCBTE Decimal(15,4),AnesthetistExcludingTax Decimal(15,4),AnesthetistIncludingTax Decimal(15,4),SelectedProducts text)
BEGIN

Insert into EventRequestsWeb
(`Event Topic`,`Event Type`,`Event Date`,`Start Time`,`End Time`,`Venue Name`,`City`,`State`,`Initiator Name`,
`Advance Amount`,`Total Expense BTC`,`Total Expense BTE`,`Total Honorarium Amount`,`Total Travel Amount`,`Total Travel & Accommodation Amount`,`Total Accommodation Amount`,
`Budget Amount`,`Total Local Conveyance`,`Total Expense`,`Initiator Email`,`RBM/BM`,`Sales Head`,`Sales Coordinator`,`Marketing Coordinator`,`Marketing Head`,
`Compliance`,`Finance Accounts`,`Finance Treasury`,`Reporting Manager`,`1 Up Manager`,`Medical Affairs Head`,`Panelists`,`Invitees`,`MIPL Invitees`,
`Brands`,`Expenses`,`SlideKits`,`BTE Expense Details`,`AttachmentPaths`,`Product Brand`,`Mode of Training`,`HOT Webinar Vendor Name`,`HOT Webinar Type`,`Venue Selection Checklist`,
`Emergency Support`,`Emergency Contact No`,`Facility Charges`,`Facility Charges BTC/BTE`,`Facility Charges Excluding Tax`,`Total Facility Charges Including Tax`,
`Anesthetist Required?`,`Anesthetist BTC/BTE`,`Anesthetist Excluding Tax`,`Anesthetist Including Tax`,`Selected Products`,`Created On`,`syncdone`,`Role`)
values
(EventTopic,EventType,EventDate,StartTime,EndTime,VenueName,City,State,InitiatorName,AdvanceAmount,TotalExpenseBTC,
TotalExpenseBTE,TotalHonorariumAmount,TotalTravelAmount,TotalTravelAccommodationAmount,TotalAccommodationAmount,BudgetAmount,TotalLocalConveyance,TotalExpense,
InitiatorEmail,RBMBM,SalesHead,SalesCoordinator,MarketingCoordinator,MarketingHead,Compliance,FinanceAccounts,FinanceTreasury,ReportingManager,
1UpManager,MedicalAffairsHead,Panelists,Invitees,MIPLInvitees,Brands,Expenses,SlideKits,BTEExpenseDetails,AttachmentPaths,ProductBrand,ModeofTraining,
HOTWebinarVendorName,HOTWebinarType,VenueSelectionChecklist,EmergencySupport,EmergencyContactNo,FacilityCharges,FacilityChargesBTCBTE,FacilityChargesExcludingTax,
TotalFacilityChargesIncludingTax,AnesthetistRequiredq,AnesthetistBTCBTE,AnesthetistExcludingTax,AnesthetistIncludingTax,SelectedProducts,now(),0,webinarRole);

SELECT LAST_INSERT_ID() as ID;
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HcpConsultantEventRequestExpensesSheet` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HcpConsultantEventRequestExpensesSheet`(Expense text,
EventIdEventRequestID varchar(20),AmountExcludingTax Decimal(15,4),
Amount Decimal(15,4),BTCBTE varchar(10),BudgetAmount Decimal(15,4),
BTCAmount Decimal(15,4),BTEAmount Decimal(15,4),
EventTopic text,EventType varchar(500),
EventDateStart date,EventEndDate date,
RegistrationAmount Decimal(15,4),RegistrationAmountExcludingTax Decimal(15,4),
VenueName text,MisCode text,RegistrationAmountBtcBte text
)
BEGIN
insert into EventRequestExpensesSheet (`Expense`,`EventId/EventRequestID`,`Amount Excluding Tax`,
`Amount`,`BTC/BTE`,`BudgetAmount`,`BTCAmount`,`BTEAmount`,`Event Topic`,`Event Type`,
`Event Date Start`,`Event End Date`,`Registration Amount`,`Registration Amount Excluding Tax`,
 `Venue name`,`MisCode`  ,`Registration Amount BTC/BTE`   )

values
(Expense ,EventIdEventRequestID,AmountExcludingTax ,
Amount ,BTCBTE ,BudgetAmount,BTCAmount ,BTEAmount ,
EventTopic ,EventType ,EventDateStart ,EventEndDate ,
RegistrationAmount ,RegistrationAmountExcludingTax ,
VenueName ,MisCode ,RegistrationAmountBtcBte );
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HCPConsultantPreEvent` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HCPConsultantPreEvent`(ApproverPreEventURL varchar(500),FinanceTreasuryURL varchar(500),InitiatorURL varchar(500),
EventTopic text,EventType varchar(100),EventDate date,EndDate date,VenueName varchar(200),
SponsorshipSocietyName varchar(200),VenueCountry varchar(200),IsAdvanceRequired varchar(50),
AdvanceAmount Decimal(15,4),Brands text,Expenses text,Panelists text,InitiatorName varchar(50),InitiatorEmail varchar(200),
RBMBM varchar(200),SalesHead varchar(200),SalesCoordinator varchar(200),MarketingCoordinator varchar(200),
MarketingHead varchar(200),Compliance varchar(200),FinanceAccounts varchar(200),FinanceTreasury text,
ReportingManager varchar(200),1UpManager varchar(200),MedicalAffairsHead text,
mRole varchar(50),TotalExpenseBTE Decimal(15,4),TotalExpenseBTC Decimal(15,4),
BTEExpenseDetails text,TotalExpense Decimal(15,4),BudgetAmount Decimal(15,4),
TotalHCPRegistrationAmount Decimal(15,4),TotalTravelAmount Decimal(15,4), TotalTravelAccommodationAmount Decimal(15,4),
 TotalAccommodationAmount Decimal(15,4), TotalLocalConveyance Decimal(15,4),
AttachmentPaths text
)
BEGIN
Insert into EventRequestsWeb
(`Approver Pre Event URL`,`Finance Treasury URL`,`Initiator URL`,`Event Topic`,`Event Type`,`Event Date`,
`End Date`,`Venue Name`,`Sponsorship Society Name`,`Venue Country`,`IsAdvanceRequired`,
`Advance Amount`,`Brands`,`Expenses`,`Panelists`,`Initiator Name`,`Initiator Email`,
`RBM/BM`,`Sales Head`,`Sales Coordinator`,`Marketing Coordinator`,`Marketing Head`,
`Compliance`,`Finance Accounts`,`Finance Treasury`,`Reporting Manager`,`1 Up Manager`,`Medical Affairs Head`,
`Role`,`Total Expense BTE`,`Total Expense BTC`,`BTE Expense Details`,`Total Expense`,
`Budget Amount`,`Total HCP Registration Amount`,`Total Travel Amount`, `Total Travel & Accommodation Amount`,
 `Total Accommodation Amount`,`Total Local Conveyance`,
`AttachmentPaths`) 
#`syncdone`,`Created On`---0 ,now()
values
(ApproverPreEventURL,FinanceTreasuryURL ,InitiatorURL,
EventTopic ,EventType ,EventDate ,EndDate ,VenueName,
SponsorshipSocietyName ,VenueCountry ,IsAdvanceRequired ,
AdvanceAmount ,Brands ,Expenses ,Panelists ,InitiatorName ,InitiatorEmail ,
RBMBM ,SalesHead ,SalesCoordinator ,MarketingCoordinator ,
MarketingHead ,Compliance ,FinanceAccounts ,FinanceTreasury ,
ReportingManager ,1UpManager ,MedicalAffairsHead ,
mRole ,TotalExpenseBTE ,TotalExpenseBTC ,
BTEExpenseDetails ,TotalExpense ,BudgetAmount ,
TotalHCPRegistrationAmount ,TotalTravelAmount , TotalTravelAccommodationAmount ,
 TotalAccommodationAmount , TotalLocalConveyance ,
AttachmentPaths  );
SELECT LAST_INSERT_ID() as ID;
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HO_EventRequestBeneficiary` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HO_EventRequestBeneficiary`(EventIdEventRequestId varchar(20),EventType varchar(500),EventDate date,VenueName varchar(500),City varchar(100),State varchar(100),
OtherCurrency varchar(100),BeneficiaryName varchar(500),BankAccountNumber varchar(100),FacilityCharges varchar(10),AnesthetistRequiredq varchar(10),Currency varchar(100),
BankName varchar(100),PANcardname varchar(500),PanNumber varchar(100),IFSCCode varchar(100),EmailId varchar(500))
BEGIN

Insert into EventRequestBeneficiary(`EventId/EventRequestId`,`EventType` ,`EventDate`,`VenueName`,`City`,`State` ,`Other Currency`,`Beneficiary Name`,
`Bank Account Number`,`Facility Charges`,`Anesthetist Required?`,`Currency`,`Bank Name`,`PAN card name`,`Pan Number`,`IFSC Code`,`Email Id`)
values
(EventIdEventRequestId,EventType,EventDate ,VenueName ,City ,State ,OtherCurrency,BeneficiaryName,BankAccountNumber,FacilityCharges,
AnesthetistRequiredq,Currency,BankName ,PANcardname ,PanNumber ,IFSCCode ,EmailId );
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HO_EventRequestHCPSlideKitDetails` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HO_EventRequestHCPSlideKitDetails`(MIS varchar(20),SlideKitType varchar(500),
SlideKitDocument text,EventIdEventRequestId varchar(20),FillersIndication text,ThreadsIndication text,
HCPName varchar(500),AttachmentPaths text
)
BEGIN
insert into EventRequestHCPSlideKitDetails(`MIS`,`Slide Kit Type`,`SlideKit Document`,`EventId/EventRequestId`,
`Fillers Indication`,`Threads Indication`,`HCP Name`,AttachmentPaths)
values
(MIS,SlideKitType,SlideKitDocument,EventIdEventRequestId,FillersIndication,ThreadsIndication,HCPName,AttachmentPaths);
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HO_EventRequestInvitees` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HO_EventRequestInvitees`(HCPName varchar(500),Designation varchar(500),EmployeeCode varchar(100),LocalConveyance varchar(10),
BTCBTE varchar(10),LcAmount Decimal(15,4),LcAmountExcludingTax Decimal(15,4),EventIdEventRequestId varchar(20),InviteeSource varchar(100),
HCPType varchar(50),MISCode varchar(100),EventTopic text,EventType text,EventDateStart date,EventEndDate date,Qualification varchar(100),Experience varchar(10))
BEGIN

insert into EventRequestInvitees(`HCPName`,`Designation`,`Employee Code`,`LocalConveyance`,`BTC/BTE`,`LcAmount`,`Lc Amount Excluding Tax`,`EventId/EventRequestId`,
`Invitee Source`,`HCP Type`,`MISCode`,`Event Topic`,`Event Type`,`Event Date Start`,`Event End Date`,`Qualification`,`Experience`)

values

(HCPName,Designation,EmployeeCode,LocalConveyance,BTCBTE,LcAmount,LcAmountExcludingTax,EventIdEventRequestId,InviteeSource,HCPType,MISCode,
EventTopic,EventType,EventDateStart,EventEndDate,Qualification,Experience);

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HO_EventRequestPanelDetails` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HO_EventRequestPanelDetails`(MISCode varchar(20),HcpRole varchar(500),HCPName varchar(500),TrainerCode varchar(50),Qualification varchar(200),Country varchar(500),
Speciality varchar(50),Tier varchar(100),HCPType varchar(50),Rationale text,FCPADate date,HonorariumRequired varchar(5),AnnualTrainerAgreementValidq varchar(10),
PresentationDuration int,PanelSessionPreparationDuration int,PanelDiscussionDuration int,QASessionDuration int,BriefingSession int,TotalSessionHours int,LCBTCBTE varchar(50),
AccomodationBTCBTE varchar(50),HonorariumAmountExcludingTax Decimal,HonorariumAmount decimal,YTDSpendIncludingCurrentEvent text,GlobalFMV varchar(10),ExpenseType varchar(500),
ModeofTravel varchar(50),TravelBTCBTE varchar(50),TravelExcludingTax decimal,Travel decimal,AccomodationExcludingTax Decimal,Accomodation Decimal,LocalConveyanceExcludingTax Decimal,
LocalConveyance Decimal,TravelAccomodationSpendIncludingCurrentEvent text,Currency varchar(50),OtherCurrency varchar(50),BeneficiaryName varchar(500),
BankAccountNumber varchar(500),BankName varchar(500),PANcardname varchar(500),PanNumber varchar(500),IFSCCode varchar(500),EmailId varchar(500),
IBNNumber varchar(100),SwiftCode varchar(100),TaxResidenceCertificate date,AgreementAmount Decimal,EventTopic text,EventType varchar(500),Venuename text,
EventDateStart date,EventEndDate date,TotalSpend varchar(50),EventIdEventRequestId varchar(20),AttachmentPaths text)
BEGIN

Insert into EventRequestPanelDetails(`MISCode`,`HcpRole`,`HCPName`,`TrainerCode`,`Qualification`,`Country`,`Speciality`,`Tier`,`HCP Type`,`Rationale`,`FCPA Date`,`HonorariumRequired`,`Annual Trainer Agreement Valid?`,`PresentationDuration`,
`PanelSessionPreparationDuration`,`PanelDiscussionDuration`,`QASessionDuration`,`BriefingSession`,`TotalSessionHours`,`LC BTC/BTE`,`Accomodation BTC/BTE`,`Honorarium Amount Excluding Tax`,
`HonorariumAmount`,`YTD Spend Including Current Event`,`Global FMV`,`ExpenseType`,`Mode of Travel`,`Travel BTC/BTE`,`Travel Excluding Tax`,
`Travel`,`Accomodation Excluding Tax`,`Accomodation`,`Local Conveyance Excluding Tax`,`LocalConveyance`,`Travel/Accomodation Spend Including Current Event`,
`Currency`,`Other Currency`,`Beneficiary Name`,`Bank Account Number`,`Bank Name`,`PAN card name`,`Pan Number`,`IFSC Code`,
`Email Id`,`IBN Number`,`Swift Code`,`Tax Residence Certificate`,`AgreementAmount`,`Event Topic`,`Event Type`,
`Venue name`,`Event Date Start`,`Event End Date`,`TotalSpend`,`EventId/EventRequestId`,AttachmentPaths)
values
(MISCode,HcpRole,HCPName,TrainerCode,Qualification,Country,Speciality,Tier,HCPType,Rationale,FCPADate,HonorariumRequired,AnnualTrainerAgreementValidq,PresentationDuration,PanelSessionPreparationDuration,
PanelDiscussionDuration,QASessionDuration,BriefingSession,TotalSessionHours,LCBTCBTE,AccomodationBTCBTE,HonorariumAmountExcludingTax,HonorariumAmount,YTDSpendIncludingCurrentEvent,
GlobalFMV,ExpenseType,ModeofTravel,TravelBTCBTE,TravelExcludingTax,Travel,AccomodationExcludingTax,Accomodation,LocalConveyanceExcludingTax,LocalConveyance,
TravelAccomodationSpendIncludingCurrentEvent,Currency,OtherCurrency,BeneficiaryName,BankAccountNumber,BankName,PANcardname,PanNumber,IFSCCode,EmailId,IBNNumber,SwiftCode,
TaxResidenceCertificate,AgreementAmount,EventTopic,EventType,Venuename,EventDateStart,EventEndDate,TotalSpend,EventIdEventRequestId,AttachmentPaths
);

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `HO_ProductBrandList` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `HO_ProductBrandList`(EventIdEventRequestID varchar(10),EventType varchar(500),EventDate date,EventTopic varchar(500),
                ProductBrand varchar(200),ProductName varchar(500),NoofSamplesRequired int)
BEGIN

Insert into EventRequestProductBrandsList(`EventId/EventRequestId`,`EventType` ,`EventDate`,`Event Topic`,`Product Brand`,`Product Name`,`No of Samples Required`)
values
(EventIdEventRequestID,EventType,EventDate,EventTopic,ProductBrand,ProductName,NoofSamplesRequired);

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `MedicalUtilityPreevent` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `MedicalUtilityPreevent`(ApproverPreEventURL varchar(500),FinanceTreasuryURL varchar(500),InitiatorURL varchar(500),
EventTopic text,EventType varchar(100),EventDate date,ValidFrom date,ValidTo date,
MedicalUtilityType varchar(200),MedicalUtilityDescription varchar(200),IsAdvanceRequired varchar(50),
AdvanceAmount Decimal(15,4),Brands text,Expenses text,Panelists text,InitiatorName varchar(50),InitiatorEmail varchar(200),
RBMBM varchar(200),SalesHead varchar(200),SalesCoordinator varchar(200),MarketingCoordinator varchar(200),
MarketingHead varchar(200),Compliance varchar(200),FinanceAccounts varchar(200),FinanceTreasury text,
ReportingManager varchar(200),1UpManager varchar(200),MedicalAffairsHead text,
mRole varchar(50),TotalExpenseBTE Decimal(15,4),TotalExpenseBTC Decimal(15,4),
BTEExpenseDetails text,TotalExpense Decimal(15,4),BudgetAmount Decimal(15,4),
AttachmentPaths text
)
BEGIN
Insert into EventRequestsWeb
(`Approver Pre Event URL`,`Finance Treasury URL`,`Initiator URL`,`Event Topic`,`Event Type`,`Event Date`,
`Valid From`,`Valid To`,`Medical Utility Type`,`Medical Utility Description`,`IsAdvanceRequired`,
`Advance Amount`,`Brands`,`Expenses`,`Panelists`,`Initiator Name`,`Initiator Email`,
`RBM/BM`,`Sales Head`,`Sales Coordinator`,`Marketing Coordinator`,`Marketing Head`,
`Compliance`,`Finance Accounts`,`Finance Treasury`,`Reporting Manager`,`1 Up Manager`,`Medical Affairs Head`,
`Role`,`Total Expense BTE`,`Total Expense BTC`,`BTE Expense Details`,`Total Expense`,
`Budget Amount`,`AttachmentPaths`) 
#`syncdone`,`Created On`---0 ,now()
values
(ApproverPreEventURL ,FinanceTreasuryURL ,InitiatorURL ,EventTopic ,EventType ,EventDate,
ValidFrom ,ValidTo,MedicalUtilityType,MedicalUtilityDescription,IsAdvanceRequired,
AdvanceAmount,Brands ,Expenses ,Panelists,InitiatorName ,InitiatorEmail ,
RBMBM,SalesHead,SalesCoordinator ,MarketingCoordinator ,MarketingHead,
Compliance ,FinanceAccounts ,FinanceTreasury,ReportingManager ,1UpManager ,MedicalAffairsHead ,
mRole,TotalExpenseBTE,TotalExpenseBTC,BTEExpenseDetails,TotalExpense,
BudgetAmount ,AttachmentPaths );
SELECT LAST_INSERT_ID() as ID;
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `MU_SPEventRequestPanelDetails` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `MU_SPEventRequestPanelDetails`(
HCPName varchar(500),MISCode varchar(20),HCPType varchar(50),Speciality varchar(50),
Tier varchar(100),MedicalUtilityCost varchar(50),MedicalUtilityType varchar(100),
MedicalUtilityDescription varchar(500),LegitimateNeed varchar(500),ObjectiveCriteria varchar(500),
Rationale text,EventTopic text, EventType varchar(500), EventDateStart date, EventEndDate date,
EventIdEventRequestId varchar(50),ExpenseType varchar(500),FCPADate date,RequestDate date,
ValidFrom date,ValidTo date,AttachmentPaths text,RegistrationAmountExcludingTax Decimal(15,4),RegistrationAmountBTCBTE varchar(10),
TravelAirBtcBte varchar(10),TravelAirIncludingTax Decimal(15,4),TravelAirExcludingTax Decimal(15,4),TravelRoadBtcBte varchar(10),
TravelRoadIncludingTax Decimal(15,4),TravelRoadExcludingTax Decimal(15,4),TravelTrainBtcBte varchar(10),TravelTrainIncludingTax Decimal(15,4),
TravelTrainExcludingTax Decimal(15,4)
)
BEGIN

insert into EventRequestPanelDetails(`HCPName`,`MISCode`,`HCP Type`,`Speciality`,`Tier`,
`Medical Utility Cost`,`Medical Utility Type`,`Medical Utility description`,`Legitimate Need`,
`Objective Criteria`,`Rationale`,`Event Topic`,`Event Type`,`Event Date Start`,
`Event End Date`,`EventId/EventRequestId`,`ExpenseType`,`FCPA Date`,`Request Date`,
`Valid From`,`Valid To`,AttachmentPaths,`Registration Amount Excluding Tax`,`Registration Amount BTC/BTE`,`TravelAirBtcBte`,`TravelAirIncludingTax`,`TravelAirExcludingTax`,
`TravelRoadBtcBte`,`TravelRoadIncludingTax`,`TravelRoadExcludingTax`,`TravelTrainBtcBte`,`TravelTrainIncludingTax`,`TravelTrainExcludingTax`)
values
(
HCPName ,MISCode ,HCPType ,Speciality,Tier,MedicalUtilityCost ,MedicalUtilityType,
MedicalUtilityDescription ,LegitimateNeed,ObjectiveCriteria ,Rationale ,EventTopic,
EventType , EventDateStart , EventEndDate,EventIdEventRequestId,ExpenseType,FCPADate,
RequestDate,ValidFrom ,ValidTo ,AttachmentPaths,RegistrationAmountExcludingTax,RegistrationAmountBTCBTE,TravelAirBtcBte,TravelAirIncludingTax,TravelAirExcludingTax,TravelRoadBtcBte,
TravelRoadIncludingTax,TravelRoadExcludingTax,TravelTrainBtcBte,TravelTrainIncludingTax,TravelTrainExcludingTax );
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `SPDeviation_Process` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `SPDeviation_Process`(EventIdEventRequestId varchar(20),EventTopic text,EventType varchar(50),EventDate date,
StartTime varchar(50),EndTime varchar(50),MISCode varchar(20),HCPName varchar(200),HonorariumAmount Decimal(15,4),
TravelAccommodationAmount Decimal(15,4),OtherExpenses Decimal(15,4),SalesHead varchar(500),
FinanceHead varchar(500),InitiatorName varchar(200),InitiatorEmail varchar(500),SalesCoordinator varchar(500),DeviationType text,
EventOpen45days varchar(10),OutstandingEvents varchar(10),EventWithin5days varchar(10),PREExpenseExcludingTax varchar(10),
TravelAccomodationExceededTrigger varchar(10),TrainerHonorariumExceededTrigger varchar(10),HCPHonorariumExceededTrigger varchar(10),
AttachmentPaths text,EndDate date,`HCPExceeds5,00,000Trigger` varchar(50) ,`HCPExceeds1,00,000Trigger`varchar(50)
)
BEGIN

insert into Deviation_Process(`EventId/EventRequestId`,`Event Topic`,`Event Type`,`Event Date`,`Start Time`,`End Time`,
`MIS Code`,`HCP Name`,`Honorarium Amount`,`Travel & Accommodation Amount`,`Other Expenses`,`Sales Head`,`Finance Head`,`Initiator Name`,
`Initiator Email`,`Sales Coordinator`,`Deviation Type`,`EventOpen45days`,`Outstanding Events`,`EventWithin5days`,`PRE-F&B Expense Excluding Tax`,
`Travel/Accomodation 3,00,000 Exceeded Trigger`,`Trainer Honorarium 12,00,000 Exceeded Trigger`,`HCP Honorarium 6,00,000 Exceeded Trigger`,AttachmentPaths,`End Date`,
`HCP exceeds 5,00,000 Trigger`,`HCP exceeds 1,00,000 Trigger`
)
values
(EventIdEventRequestId,EventTopic,EventType,EventDate,StartTime,EndTime,MISCode,HCPName,HonorariumAmount,TravelAccommodationAmount,OtherExpenses,
SalesHead,FinanceHead,InitiatorName,InitiatorEmail,SalesCoordinator,DeviationType,EventOpen45days,OutstandingEvents,EventWithin5days,
PREExpenseExcludingTax,TravelAccomodationExceededTrigger,TrainerHonorariumExceededTrigger,HCPHonorariumExceededTrigger,AttachmentPaths,EndDate,
`HCPExceeds5,00,000Trigger` ,`HCPExceeds1,00,000Trigger`
);

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `SPEventRequestExpensesSheet` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `SPEventRequestExpensesSheet`(Expense text,EventIdEventRequestID varchar(20),AmountExcludingTaxq Decimal(15,4),AmountExcludingTax Decimal(15,4),
Amount Decimal(15,4),BTCBTE varchar(10),BudgetAmount Decimal(15,4),BTCAmount Decimal(15,4),BTEAmount Decimal(15,4),EventTopic text,EventType varchar(500),
EventDateStart date,EventEndDate date)
BEGIN
insert into EventRequestExpensesSheet (`Expense`,`EventId/EventRequestID`,`AmountExcludingTax?`,`Amount Excluding Tax`,`Amount`,
`BTC/BTE`,`BudgetAmount`,`BTCAmount`,`BTEAmount`,`Event Topic`,`Event Type`,`Event Date Start`,`Event End Date`)
values
(Expense,EventIdEventRequestID,AmountExcludingTaxq,AmountExcludingTax,Amount,BTCBTE,BudgetAmount,BTCAmount,BTEAmount,EventTopic,
EventType,EventDateStart,EventEndDate);
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `SPEventRequestHCPSlideKitDetails` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `SPEventRequestHCPSlideKitDetails`(MIS varchar(20),SlideKitType varchar(500),
SlideKitDocument text,EventIdEventRequestId varchar(20),AttachmentPaths text
)
BEGIN
insert into EventRequestHCPSlideKitDetails(`MIS`,`Slide Kit Type`,`SlideKit Document`,`EventId/EventRequestId`,AttachmentPaths)
values
(MIS,SlideKitType,SlideKitDocument,EventIdEventRequestId,AttachmentPaths);
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `SPEventRequestInvitees` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `SPEventRequestInvitees`(HCPName varchar(500),Designation varchar(500),EmployeeCode varchar(100),LocalConveyance varchar(10),
BTCBTE varchar(10),LcAmount Decimal(15,4),LcAmountExcludingTax Decimal(15,4),EventIdEventRequestId varchar(20),InviteeSource varchar(100),
HCPType varchar(50),Speciality varchar(50),MISCode varchar(100),EventTopic text,EventType text,EventDateStart date,EventEndDate date)
BEGIN


insert into EventRequestInvitees(`HCPName`,`Designation`,`Employee Code`,`LocalConveyance`,`BTC/BTE`,`LcAmount`,`Lc Amount Excluding Tax`,`EventId/EventRequestId`,
`Invitee Source`,`HCP Type`,`Speciality`,`MISCode`,`Event Topic`,`Event Type`,`Event Date Start`,`Event End Date`)
values
(HCPName,Designation,EmployeeCode,LocalConveyance,BTCBTE,LcAmount,LcAmountExcludingTax,EventIdEventRequestId,InviteeSource,HCPType,Speciality,MISCode,
EventTopic,EventType,EventDateStart,EventEndDate);

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `SPEventRequestPanelDetails` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `SPEventRequestPanelDetails`(`HcpRole`varchar(500),`MISCode`varchar(20),`Travel`Decimal(15,4),`TotalSpend`varchar(50),`Accomodation`Decimal(15,4),
`LocalConveyance`Decimal(15,4),`SpeakerCode`varchar(50),`TrainerCode` varchar(50),`HonorariumRequired` varchar(5),`AgreementAmount`Decimal(15,4),`HonorariumAmount`Decimal(15,4),
`Speciality`varchar(50),`EventTopic`text,`EventType`varchar(500),`EventDateStart`date,`EventEndDate`date,`HCPName`varchar(500),`PANcardname`varchar(500),
`ExpenseType`varchar(500),`BankAccountNumber`varchar(500),`BankName`varchar(500),`IFSCCode`varchar(500),`FCPADate`date,`Currency`varchar(50),`HonorariumAmountExcludingTax`Decimal(15,4),
`TravelExcludingTax`Decimal(15,4),`AccomodationExcludingTax`Decimal(15,4),`LocalConveyanceExcludingTax`Decimal(15,4),`LCBTCBTE`varchar(50),`TravelBTCBTE`varchar(50),
`AccomodationBTCBTE`varchar(50),`ModeofTravel`varchar(50),`OtherCurrency`varchar(50),`BeneficiaryName`varchar(500),`PanNumber`varchar(500),`OtherType`text,
`Tier`varchar(100),`HCPType`varchar(50),`PresentationDuration`varchar(10),`PanelSessionPreparationDuration`varchar(10),`PanelDiscussionDuration`varchar(10),
`QASessionDuration`varchar(10),`BriefingSession`varchar(10),`TotalSessionHours`varchar(10),`Rationale`text,`EventIdEventRequestId`varchar(50),AttachmentPaths text
)
BEGIN

insert into EventRequestPanelDetails(`HcpRole`,`MISCode`,`Travel`,`TotalSpend`,`Accomodation`,`LocalConveyance`,`SpeakerCode`,`TrainerCode`,`HonorariumRequired`,
`AgreementAmount`,`HonorariumAmount`,`Speciality`,`Event Topic`,`Event Type`,`Event Date Start`,`Event End Date`,`HCPName`,`PAN card name`,`ExpenseType`,
`Bank Account Number`,`Bank Name`,`IFSC Code`,`FCPA Date`,`Currency`,`Honorarium Amount Excluding Tax`,`Travel Excluding Tax`,`Accomodation Excluding Tax`,
`Local Conveyance Excluding Tax`,`LC BTC/BTE`,`Travel BTC/BTE`,`Accomodation BTC/BTE`,`Mode of Travel`,`Other Currency`,`Beneficiary Name`,`Pan Number`,
`Other Type`,`Tier`,`HCP Type`,`PresentationDuration`,`PanelSessionPreparationDuration`,`PanelDiscussionDuration`,`QASessionDuration`,
`BriefingSession`,`TotalSessionHours`,`Rationale `,`EventId/EventRequestId`,AttachmentPaths)
values
(HcpRole,MISCode,Travel,TotalSpend,Accomodation,LocalConveyance,SpeakerCode,TrainerCode,HonorariumRequired,AgreementAmount,HonorariumAmount,Speciality,
EventTopic,EventType,EventDateStart,EventEndDate,HCPName,PANcardname,ExpenseType,BankAccountNumber,BankName,IFSCCode,FCPADate,Currency,
HonorariumAmountExcludingTax,TravelExcludingTax,AccomodationExcludingTax,LocalConveyanceExcludingTax,LCBTCBTE,TravelBTCBTE,AccomodationBTCBTE,
ModeofTravel,OtherCurrency,BeneficiaryName,PanNumber,OtherType,Tier,HCPType,PresentationDuration,PanelSessionPreparationDuration,
PanelDiscussionDuration,QASessionDuration,BriefingSession,TotalSessionHours,Rationale ,EventIdEventRequestId,AttachmentPaths
);

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `SPEventRequestsBrandsList` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `SPEventRequestsBrandsList`(Allocation varchar(50),Brands varchar(100),
ProjectID varchar(100),EventIdEventRequestId varchar(50)
)
BEGIN
Insert into EventRequestsBrandsList(`% Allocation`,`Brands`,`Project ID`,`EventId/EventRequestId`)
values
(Allocation,Brands,ProjectID,EventIdEventRequestId);
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `StallFabricationPreevent` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `StallFabricationPreevent`(ApproverPreEventURL varchar(500),FinanceTreasuryURL varchar(500),InitiatorURL varchar(500),
EventTopic text,EventType varchar(100),EventDate date,EndDate date,
Class_III_EventCode varchar(50),RBMBM varchar(200),SalesHead varchar(200),MedicalAffairsHead text,
ReportingManager varchar(200),1UpManager varchar(200),Compliance varchar(200),
FinanceAccounts varchar(200),SalesCoordinator varchar(200),MarketingCoordinator varchar(200),
FinanceTreasury text,InitiatorName varchar(50),InitiatorEmail varchar(200),BudgetAmount Decimal(15,4),
IsAdvanceRequired varchar(200),
AdvanceAmount Decimal(15,4),TotalExpenseBTE Decimal(15,4),TotalExpenseBTC Decimal(15,4),AttachmentPaths text,
BTEExpenseDetails text,
webinarRole varchar(50),Brands text,Expenses text,TotalExpense Decimal(15,4),MarketingHead varchar(200)

)
BEGIN
Insert into EventRequestsWeb
(`Approver Pre Event URL`,`Finance Treasury URL`,`Initiator URL`,`Event Topic`,`Event Type`,`Event Date`,`End Date`,`Class III Event Code`,`RBM/BM`,`Sales Head`,
`Medical Affairs Head`,`Reporting Manager`,`1 Up Manager`,`Compliance`,`Finance Accounts`,`Sales Coordinator`,`Marketing Coordinator`,`Finance Treasury`,
`Initiator Name`,`Initiator Email`,`Budget Amount`,`IsAdvanceRequired`,`Advance Amount`,`Total Expense BTE`,`Total Expense BTC`,
`AttachmentPaths`,`BTE Expense Details`,`Role`,`Brands`,`Expenses`,`Total Expense`,`Marketing Head`,`syncdone`,`Created On`) 
values
(ApproverPreEventURL ,FinanceTreasuryURL ,InitiatorURL ,EventTopic ,EventType ,EventDate ,EndDate,Class_III_EventCode,RBMBM,
SalesHead,MedicalAffairsHead ,ReportingManager ,1UpManager ,Compliance ,FinanceAccounts ,SalesCoordinator ,MarketingCoordinator ,
FinanceTreasury,InitiatorName ,InitiatorEmail ,BudgetAmount ,IsAdvanceRequired ,AdvanceAmount,TotalExpenseBTE,TotalExpenseBTC,
AttachmentPaths,BTEExpenseDetails,webinarRole,Brands ,Expenses ,TotalExpense,MarketingHead ,0 ,now());
SELECT LAST_INSERT_ID() as ID;
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `UpdateEventid` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `UpdateEventid`(Dbid bigint,Eventid varchar(50))
BEGIN

update EventRequestsWeb set `EventId/EventRequestId` = Eventid,syncdone=1,Modified = now() where ID = Dbid;

END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 DROP PROCEDURE IF EXISTS `WebinarPreevent` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_0900_ai_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
CREATE DEFINER=`menarini`@`%` PROCEDURE `WebinarPreevent`(ApproverPreEventURL varchar(500),FinanceTreasuryURL varchar(500),InitiatorURL varchar(500),
EventTopic text,EventType varchar(100),EventDate date,StartTime varchar(15),EndTime varchar(15),MeetingType text,Brands text,Expenses text,
Panelists text,Invitees text,MIPLInvitees text,SlideKits text,IsAdvanceRequired varchar(5),EventOpen30days varchar(5),EventWithin7days varchar(5),
InitiatorName varchar(100),AdvanceAmount Decimal(15,4),TotalExpenseBTC Decimal(15,4),TotalExpenseBTE Decimal(15,4),TotalHonorariumAmount Decimal(15,4),
TotalTravelAmount Decimal(15,4),TotalTravelAccommodationAmount Decimal(15,4),TotalAccommodationAmount Decimal(15,4),BudgetAmount Decimal(15,4),TotalLocalConveyance Decimal(15,4),
TotalExpense Decimal(15,4),InitiatorEmail varchar(200),RBMBM varchar(200),SalesHead varchar(200),SalesCoordinator varchar(200),MarketingCoordinator varchar(200),
MarketingHead varchar(200),Compliance varchar(200),FinanceAccounts varchar(200),FinanceTreasury varchar(200),ReportingManager varchar(200),
1UpManager varchar(200),MedicalAffairsHead varchar(200),BTEExpenseDetails text,AttachmentPaths text,webinarRole varchar(50)
)
BEGIN

Insert into EventRequestsWeb
(`Approver Pre Event URL`,`Finance Treasury URL`,`Initiator URL`,`Event Topic`,`Event Type`,`Event Date`,`Start Time`,`End Time`,
`Meeting Type`,`Brands`,`Expenses`,`Panelists`,`Invitees`,`MIPL Invitees`,`SlideKits`,`IsAdvanceRequired`,`EventOpen30days`,`EventWithin7days`,`Initiator Name`,
`Advance Amount`,`Total Expense BTC`,`Total Expense BTE`,`Total Honorarium Amount`,`Total Travel Amount`,`Total Travel & Accommodation Amount`,`Total Accommodation Amount`,
`Budget Amount`,`Total Local Conveyance`,`Total Expense`,`Initiator Email`,`RBM/BM`,`Sales Head`,`Sales Coordinator`,`Marketing Coordinator`,
`Marketing Head`,`Compliance`,`Finance Accounts`,`Finance Treasury`,`Reporting Manager`,`1 Up Manager`,`Medical Affairs Head`,`BTE Expense Details`,`AttachmentPaths`,`Created On`,`syncdone`,`Role`) 
values
(ApproverPreEventURL,FinanceTreasuryURL,InitiatorURL,EventTopic,EventType,EventDate,StartTime,EndTime,MeetingType,
Brands,Expenses,Panelists,Invitees,MIPLInvitees,SlideKits,IsAdvanceRequired,EventOpen30days,EventWithin7days,InitiatorName,AdvanceAmount,TotalExpenseBTC,
TotalExpenseBTE,TotalHonorariumAmount,TotalTravelAmount,TotalTravelAccommodationAmount,TotalAccommodationAmount,BudgetAmount,TotalLocalConveyance,TotalExpense,
InitiatorEmail,RBMBM,SalesHead,SalesCoordinator,MarketingCoordinator,MarketingHead,Compliance,FinanceAccounts,FinanceTreasury,ReportingManager,
1UpManager,MedicalAffairsHead,BTEExpenseDetails,AttachmentPaths,now(),0,webinarRole);
SELECT LAST_INSERT_ID() as ID;
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2024-08-08 12:40:14
