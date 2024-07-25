CREATE DATABASE  IF NOT EXISTS `vehiculos` /*!40100 DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci */ /*!80016 DEFAULT ENCRYPTION='N' */;
USE `vehiculos`;


CREATE TABLE `vehiculos`.`estadosvehiculares` (
  `id` INT NOT NULL AUTO_INCREMENT,
  `plataforma` VARCHAR(10) NULL,
  `estado` VARCHAR(15) NULL DEFAULT 'No ejecutado',
  PRIMARY KEY (`id`)
);

INSERT INTO `vehiculos`.`estadosvehiculares` (`id`, `plataforma`, `estado`)
VALUES
(1, 'Ituran', 'No ejecutado'),
(2, 'MDVR', 'No ejecutado'),
(3, 'Ubicar', 'No ejecutado'),
(4, 'Ubicom', 'No ejecutado'),
(5, 'Securitrac', 'No ejecutado'),
(6, 'Wialon', 'No ejecutado');

CREATE TABLE `vehiculos`.`tablaerrores` (
  `id` INT NOT NULL AUTO_INCREMENT,
  `plataforma` VARCHAR(16),
  `fecha` DATETIME,
  `estado` VARCHAR(16),
  PRIMARY KEY (`id`)
);

CREATE TABLE `vehiculos`.`fueralaboral` (
  `id` INT NOT NULL AUTO_INCREMENT,
  `placa` VARCHAR(8),
  `fecha` DATETIME,
  PRIMARY KEY (`id`)
);

CREATE TABLE `vehiculos`.`plataformasvehiculares` (
  id INT NOT NULL AUTO_INCREMENT,
  `correo` VARCHAR(64),
  `correoCopia` VARCHAR(64),
  PRIMARY KEY (`id`)
);

INSERT INTO `vehiculos`.`plataformasvehiculares` (`id`, `correo`, `correocopia`)
VALUES
(1, 'directorhseq.laboratorio@sgiltda.com', 'sophya.viveros@sgiltda.com'),
(2, '', 'desarrollo.software@sgiltda.com');
