-- Script d'import pour SQL Server
-- Genere le 02/04/2026 21:13:37
-- Total: 29 enregistrements

USE [MergeDB];
GO

-- Creation de la table (si elle n'existe pas)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ref_obj_gestion')
BEGIN
    CREATE TABLE [ref_obj_gestion] (
        [Id] INT IDENTITY(1,1) PRIMARY KEY,
        [Date] DATETIME NULL,
        [Used] NVARCHAR(255) NULL,
        [NomFr] NVARCHAR(500) NULL,
        [Definition] NVARCHAR(MAX) NULL,
        [NomEn] NVARCHAR(500) NULL,
        [Trigramme] NVARCHAR(100) NULL,
        [NomFichierExcel] NVARCHAR(500) NULL,
        [DateImport] DATETIME DEFAULT GETDATE()
    );
    CREATE INDEX IX_ref_obj_gestion_NomFr ON [ref_obj_gestion] ([NomFr]);
    CREATE INDEX IX_ref_obj_gestion_Trigramme ON [ref_obj_gestion] ([Trigramme]);
    CREATE INDEX IX_ref_obj_gestion_NomFichierExcel ON [ref_obj_gestion] ([NomFichierExcel]);
END
GO

-- Vider la table avant insertion
TRUNCATE TABLE [ref_obj_gestion];
GO

-- Insertion des donnees
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'1', NULL, N'1', N'1', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'2', NULL, N'2', N'2', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'3', NULL, N'3', N'3', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'4', NULL, N'4', N'4', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'5', NULL, N'5', N'5', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'6', NULL, N'6', N'6', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'7', NULL, N'7', N'7', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'8', NULL, N'8', N'8', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-10 00:00:00', NULL, N'9', NULL, N'9', N'9', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-02-07 00:00:00', NULL, N'"a l''autorité parentale"', NULL, N'"got the parental authority"', N'GOTPAAAHR', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-02-07 00:00:00', NULL, N'"a le pouvoir bancaire"', NULL, N'"got the banking power"', N'GOTBNKPOW', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-02-06 00:00:00', NULL, N'"a une entreprise individuelle"', NULL, N'"got an individual company"', N'GOTIDVCNY', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-02-07 00:00:00', NULL, N'"appartient au foyer fiscal"', NULL, N'"belonging to the fiscal household"', N'BLGFSCHHD', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-02-07 00:00:00', NULL, N'"bénéficiaire effectif"', NULL, N'"effective beneficiary"', N'EFTBEN', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2018-10-19 00:00:00', NULL, N'"Centre Est Europe"', NULL, N'"Centre Est Europe"', N'CEE', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-11-12 00:00:00', NULL, N'2nd meilleur', NULL, NULL, N'TO2', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-07-06 00:00:00', NULL, N'36 derniers mois', NULL, N'last 36 months', N'LST36M', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-01-03 00:00:00', NULL, N'à payer', NULL, N'required to pay', N'RTP', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-01-03 00:00:00', NULL, N'abonné', NULL, N'subscriber', N'SBC', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2021-01-03 00:00:00', NULL, N'absence', NULL, N'absence', N'ABS', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2019-09-16 00:00:00', NULL, N'ACAM', N'Autorité de Contrôle des Assurances et des Mutuelles', N'ACAM', N'AAM', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-09-08 00:00:00', NULL, N'accès', NULL, N'access', N'AKS', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-09-08 00:00:00', NULL, N'accès confidentiel', NULL, N'confidential access', N'CFDAKS', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-07-28 00:00:00', NULL, N'accident', NULL, N'accident', N'ACD', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-08-11 00:00:00', NULL, N'accident de circulation', NULL, N'Traffic accident', N'TFFACD', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-06-30 00:00:00', NULL, N'accord de gestion', NULL, N'management agreement', N'MGTAGM', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2020-08-06 00:00:00', NULL, N'accouchement', NULL, N'delivery', N'DLY', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2019-09-11 00:00:00', NULL, N'ACEM', N'Action Clé En Main', N'turnkey action', N'TNKACT', N'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO [ref_obj_gestion] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])
VALUES ('2019-09-13 00:00:00', NULL, N'ACEM réalisée', NULL, N'turnkey action completion', N'TNKACTCPI', N'IBIA_ACCMGR_Outil.xlsx');
GO

-- Verification
SELECT COUNT(*) AS total_records FROM [ref_obj_gestion];
GO
