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
-- Table structure for table `EventRequestProductBrandsList`
--

DROP TABLE IF EXISTS `EventRequestProductBrandsList`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `EventRequestProductBrandsList` (
  `Id` int NOT NULL AUTO_INCREMENT,
  `Product ID` varchar(15) DEFAULT NULL,
  `EventId/EventRequestId` varchar(20) DEFAULT NULL,
  `EventType` varchar(500) DEFAULT NULL,
  `EventDate` date DEFAULT NULL,
  `Event Topic` varchar(500) DEFAULT NULL,
  `Product Brand` varchar(200) DEFAULT NULL,
  `Product Name` varchar(500) DEFAULT NULL,
  `No of Samples Required` int DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=18 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `EventRequestProductBrandsList`
--

LOCK TABLES `EventRequestProductBrandsList` WRITE;
/*!40000 ALTER TABLE `EventRequestProductBrandsList` DISABLE KEYS */;
INSERT INTO `EventRequestProductBrandsList` (`Id`, `Product ID`, `EventId/EventRequestId`, `EventType`, `EventDate`, `Event Topic`, `Product Brand`, `Product Name`, `No of Samples Required`) VALUES (1,NULL,'275','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse touch  Filler  + Lidocaine',4),(2,NULL,'275','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse restore  Filler  + Lidocaine',4),(3,NULL,'275','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse threads','Definisse Nose Threads',6),(4,NULL,'276','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse touch  Filler  + Lidocaine',4),(5,NULL,'276','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse restore  Filler  + Lidocaine',4),(6,NULL,'276','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse threads','Definisse Nose Threads',6),(7,NULL,'282','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse touch  Filler  + Lidocaine',4),(8,NULL,'282','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse restore  Filler  + Lidocaine',4),(9,NULL,'282','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse threads','Definisse Nose Threads',6),(10,NULL,'286','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse touch  Filler  + Lidocaine',4),(11,NULL,'286','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse restore  Filler  + Lidocaine',4),(12,NULL,'286','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse threads','Definisse Nose Threads',6),(13,NULL,'290','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse touch  Filler  + Lidocaine',4),(14,NULL,'290','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse fillers','Definisse restore  Filler  + Lidocaine',4),(15,NULL,'290','Hands on Training Workshops','2024-07-04','Hands On Payload Test','Definisse threads','Definisse Nose Threads',6),(16,NULL,'293','Hands on Training Workshops','2024-07-18','Test','Definisse fillers','Definisse restore  Filler  + Lidocaine',1),(17,NULL,'293','Hands on Training Workshops','2024-07-18','Test','Definisse threads','Definisse Free Floating Threads 23cm',3);
/*!40000 ALTER TABLE `EventRequestProductBrandsList` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2024-08-08 12:40:06
