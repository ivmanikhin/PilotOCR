-- --------------------------------------------------------
-- Хост:                         127.0.0.1
-- Версия сервера:               10.7.3-MariaDB - mariadb.org binary distribution
-- Операционная система:         Win64
-- HeidiSQL Версия:              11.3.0.6295
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;


-- Дамп структуры базы данных pilotsql
CREATE DATABASE IF NOT EXISTS `pilotsql` /*!40100 DEFAULT CHARACTER SET utf8mb3 */;
USE `pilotsql`;

-- Дамп структуры для таблица pilotsql.inbox
CREATE TABLE IF NOT EXISTS `inbox` (
  `letter_counter` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `out_no` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `doc_id` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `date` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `subject` text CHARACTER SET utf8mb3 DEFAULT NULL,
  `correspondent` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `text` mediumtext CHARACTER SET utf8mb3 DEFAULT NULL,
  `year` year(4) DEFAULT NULL,
  `unrecognized` text CHARACTER SET utf8mb3 DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- Экспортируемые данные не выделены.

-- Дамп структуры для таблица pilotsql.sent
CREATE TABLE IF NOT EXISTS `sent` (
  `letter_counter` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `out_no` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `doc_id` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `date` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `subject` text CHARACTER SET utf8mb3 DEFAULT NULL,
  `correspondent` tinytext CHARACTER SET utf8mb3 DEFAULT NULL,
  `text` mediumtext CHARACTER SET utf8mb3 DEFAULT NULL,
  `year` year(4) DEFAULT NULL,
  `unrecognized` text CHARACTER SET utf8mb3 DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

-- Экспортируемые данные не выделены.

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IFNULL(@OLD_FOREIGN_KEY_CHECKS, 1) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40111 SET SQL_NOTES=IFNULL(@OLD_SQL_NOTES, 1) */;
