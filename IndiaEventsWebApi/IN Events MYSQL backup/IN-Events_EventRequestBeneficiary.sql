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
-- Table structure for table `EventRequestBeneficiary`
--

DROP TABLE IF EXISTS `EventRequestBeneficiary`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `EventRequestBeneficiary` (
  `Id` int NOT NULL AUTO_INCREMENT,
  `Benf Id` varchar(20) DEFAULT NULL,
  `EventId/EventRequestId` varchar(20) DEFAULT NULL,
  `EventType` varchar(500) DEFAULT NULL,
  `EventDate` date DEFAULT NULL,
  `VenueName` varchar(500) DEFAULT NULL,
  `State` varchar(100) DEFAULT NULL,
  `City` varchar(100) DEFAULT NULL,
  `Facility Charges` varchar(10) DEFAULT NULL,
  `Anesthetist Required?` varchar(10) DEFAULT NULL,
  `Type of Beneficiary` varchar(100) DEFAULT NULL,
  `Currency` varchar(100) DEFAULT NULL,
  `Other Currency` varchar(100) DEFAULT NULL,
  `Beneficiary Name` varchar(500) DEFAULT NULL,
  `Bank Account Number` varchar(100) DEFAULT NULL,
  `Bank Name` varchar(100) DEFAULT NULL,
  `PAN card name` varchar(500) DEFAULT NULL,
  `Pan Number` varchar(100) DEFAULT NULL,
  `IFSC Code` varchar(100) DEFAULT NULL,
  `Email Id` varchar(500) DEFAULT NULL,
  `Swift Code` varchar(100) DEFAULT NULL,
  `IBN Number` varchar(100) DEFAULT NULL,
  `Tax Residence Certificate Date` date DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `EventRequestBeneficiary`
--

LOCK TABLES `EventRequestBeneficiary` WRITE;
/*!40000 ALTER TABLE `EventRequestBeneficiary` DISABLE KEYS */;
INSERT INTO `EventRequestBeneficiary` (`Id`, `Benf Id`, `EventId/EventRequestId`, `EventType`, `EventDate`, `VenueName`, `State`, `City`, `Facility Charges`, `Anesthetist Required?`, `Type of Beneficiary`, `Currency`, `Other Currency`, `Beneficiary Name`, `Bank Account Number`, `Bank Name`, `PAN card name`, `Pan Number`, `IFSC Code`, `Email Id`, `Swift Code`, `IBN Number`, `Tax Residence Certificate Date`) VALUES (1,NULL,'275','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(2,NULL,'276','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(3,NULL,'277','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(4,NULL,'279','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(5,NULL,'280','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(6,NULL,'281','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(7,NULL,'282','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(8,NULL,'286','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL),(9,NULL,'290','Hands on Training Workshops','2024-07-04','Area 51','Uttarakhand','Kashipur','No','Yes',NULL,'INR',' ','APLESH KHICHADIYA','50100199198670','ANZ7878','APLESH KHICHADIYA','SDEW121313',' ',' ',NULL,NULL,NULL);
/*!40000 ALTER TABLE `EventRequestBeneficiary` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2024-08-08 12:39:45
