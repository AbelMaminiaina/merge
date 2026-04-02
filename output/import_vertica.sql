-- Script d'import pour Vertica
-- Genere le 02/04/2026 02:12:38
-- Total: 29 enregistrements

-- Creation de la table (si elle n'existe pas)
CREATE TABLE IF NOT EXISTS ref_obj_gestion (
    Id IDENTITY(1,1),
    Date TIMESTAMP NULL,
    Used VARCHAR(255),
    NomFr VARCHAR(500),
    Definition VARCHAR(65000),
    NomEn VARCHAR(500),
    Trigramme VARCHAR(100),
    NomFichierExcel VARCHAR(500),
    DateImport TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Vider la table avant insertion
TRUNCATE TABLE ref_obj_gestion;

-- Insertion des donnees
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '1', NULL, '1', '1', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '2', NULL, '2', '2', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '3', NULL, '3', '3', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '4', NULL, '4', '4', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '5', NULL, '5', '5', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '6', NULL, '6', '6', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '7', NULL, '7', '7', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '8', NULL, '8', '8', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-10 00:00:00', NULL, '9', NULL, '9', '9', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-02-07 00:00:00', NULL, '"a l''autorité parentale"', NULL, '"got the parental authority"', 'GOTPAAAHR', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-02-07 00:00:00', NULL, '"a le pouvoir bancaire"', NULL, '"got the banking power"', 'GOTBNKPOW', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-02-06 00:00:00', NULL, '"a une entreprise individuelle"', NULL, '"got an individual company"', 'GOTIDVCNY', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-02-07 00:00:00', NULL, '"appartient au foyer fiscal"', NULL, '"belonging to the fiscal household"', 'BLGFSCHHD', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-02-07 00:00:00', NULL, '"bénéficiaire effectif"', NULL, '"effective beneficiary"', 'EFTBEN', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2018-10-19 00:00:00', NULL, '"Centre Est Europe"', NULL, '"Centre Est Europe"', 'CEE', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-11-12 00:00:00', NULL, '2nd meilleur', NULL, NULL, 'TO2', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-07-06 00:00:00', NULL, '36 derniers mois', NULL, 'last 36 months', 'LST36M', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-01-03 00:00:00', NULL, 'à payer', NULL, 'required to pay', 'RTP', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-01-03 00:00:00', NULL, 'abonné', NULL, 'subscriber', 'SBC', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2021-01-03 00:00:00', NULL, 'absence', NULL, 'absence', 'ABS', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2019-09-16 00:00:00', NULL, 'ACAM', 'Autorité de Contrôle des Assurances et des Mutuelles', 'ACAM', 'AAM', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-09-08 00:00:00', NULL, 'accès', NULL, 'access', 'AKS', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-09-08 00:00:00', NULL, 'accès confidentiel', NULL, 'confidential access', 'CFDAKS', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-07-28 00:00:00', NULL, 'accident', NULL, 'accident', 'ACD', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-08-11 00:00:00', NULL, 'accident de circulation', NULL, 'Traffic accident', 'TFFACD', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-06-30 00:00:00', NULL, 'accord de gestion', NULL, 'management agreement', 'MGTAGM', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2020-08-06 00:00:00', NULL, 'accouchement', NULL, 'delivery', 'DLY', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2019-09-11 00:00:00', NULL, 'ACEM', 'Action Clé En Main', 'turnkey action', 'TNKACT', 'IBIA_ACCMGR_Outil.xlsx');
INSERT INTO ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
VALUES ('2019-09-13 00:00:00', NULL, 'ACEM réalisée', NULL, 'turnkey action completion', 'TNKACTCPI', 'IBIA_ACCMGR_Outil.xlsx');

COMMIT;

-- Verification
SELECT COUNT(*) AS total_records FROM ref_obj_gestion;
