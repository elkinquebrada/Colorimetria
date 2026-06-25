-- 0. Creación de la base de datos (Ejecutar primero)
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'ColorimetriaDB')
BEGIN
    CREATE DATABASE ColorimetriaDB;
END
GO

USE ColorimetriaDB;
GO

-- 1. Creación de la tabla de cabecera
-- Almacena los datos globales del lote, matiz y las lecturas instrumentales de los iluminantes.
CREATE TABLE tbl_analisis_cabecera (
    Id_Lote INT PRIMARY KEY IDENTITY(1,1),
    ShadeName VARCHAR(50) NOT NULL,
    LotNo VARCHAR(50) NOT NULL,
    FechaRegistro DATETIME DEFAULT GETDATE(),
    DeltaE_TL84 DECIMAL(8,4) NULL,
    CMC_TL84 DECIMAL(8,4) NULL,
    Status_TL84 VARCHAR(10) NULL,
    DeltaE_A DECIMAL(8,4) NULL,
    CMC_A DECIMAL(8,4) NULL,
    Status_A VARCHAR(10) NULL
);

-- 2. Creación de la tabla de detalle
-- Contiene la formulación original y las tres recetas derivadas (Concentraciones, Proporciones y Ajustes).
CREATE TABLE tbl_analisis_detalle (
    Id_Detalle INT PRIMARY KEY IDENTITY(1,1),
    Id_Lote INT NOT NULL,
    DyeCode VARCHAR(30) NOT NULL,
    DyeName VARCHAR(100) NOT NULL,
    Concentration_Original DECIMAL(18,5) NOT NULL,
    Proportion_Original DECIMAL(5,2) NOT NULL,
    
    -- Receta 1 (Luminosidad)
    R1_Con_Percentage DECIMAL(18,5) NULL,
    R1_Part_Percentage DECIMAL(5,2) NULL,
    R1_Ajuste_Percentage DECIMAL(5,2) NULL,
    
    -- Receta 2 (Croma)
    R2_Con_Percentage DECIMAL(18,5) NULL,
    R2_Part_Percentage DECIMAL(5,2) NULL,
    R2_Ajuste_Percentage DECIMAL(5,2) NULL,
    
    -- Receta 3 (Tono)
    R3_Con_Percentage DECIMAL(18,5) NULL,
    R3_Part_Percentage DECIMAL(5,2) NULL,
    R3_Ajuste_Percentage DECIMAL(5,2) NULL,
    
    CONSTRAINT FK_Analisis_Detalle_Cabecera FOREIGN KEY (Id_Lote) 
        REFERENCES tbl_analisis_cabecera(Id_Lote) ON DELETE CASCADE
);

-- 3. Índices de rendimiento
CREATE INDEX IX_Cabecera_Shade_Lot ON tbl_analisis_cabecera(ShadeName, LotNo);
CREATE INDEX IX_Detalle_IdLote ON tbl_analisis_detalle(Id_Lote);
