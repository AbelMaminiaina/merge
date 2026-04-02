-- Script d'import pour Vertica (optimise pour gros volumes)
-- Genere le 02/04/2026 22:16:47
-- Total: 29 enregistrements
-- Methode: COPY FROM LOCAL (beaucoup plus rapide que INSERT)

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

-- Chargement des donnees depuis le fichier CSV
-- IMPORTANT: Le fichier CSV doit etre dans le meme dossier que ce script
COPY ref_obj_gestion (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)
FROM LOCAL 'C:\Users\amami\GitHub\merge\output\import_vertica_data.csv'
DELIMITER ','
ENCLOSED BY '"'
SKIP 1
NULL ''
REJECTED DATA AS TABLE ref_obj_gestion_rejects
EXCEPTIONS AS TABLE ref_obj_gestion_exceptions;

COMMIT;

-- Verification
SELECT COUNT(*) AS total_records FROM ref_obj_gestion;

-- Verifier les rejets (si erreurs)
-- SELECT * FROM ref_obj_gestion_rejects;
