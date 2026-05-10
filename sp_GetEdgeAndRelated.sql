-- ============================================================
-- Procédure stockée : sp_GetEdgeAndRelated
-- Table             : edge (key, data1, data2, trans, output, code_ope)
-- Description       : Retourne l'edge source + les edges liés
--                     (via data1/data2), limités à @top lignes,
--                     avec le total global en colonne TotalLignes.
-- ============================================================

-- ------------------------------------------------------------
-- 1. INDEX RECOMMANDÉS (à exécuter une seule fois)
-- ------------------------------------------------------------

-- Index cluster sur la PK (à créer si inexistant)
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('edge')
      AND name = 'CIX_edge_key'
)
BEGIN
    CREATE UNIQUE CLUSTERED INDEX CIX_edge_key
        ON edge ([key]);
END;
GO

-- Index couvrant sur data1
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('edge')
      AND name = 'NIX_edge_data1'
)
BEGIN
    CREATE NONCLUSTERED INDEX NIX_edge_data1
        ON edge (data1)
        INCLUDE ([key], data2, trans, output, code_ope);
END;
GO

-- Index couvrant sur data2
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('edge')
      AND name = 'NIX_edge_data2'
)
BEGIN
    CREATE NONCLUSTERED INDEX NIX_edge_data2
        ON edge (data2)
        INCLUDE ([key], data1, trans, output, code_ope);
END;
GO

-- ------------------------------------------------------------
-- 2. PROCÉDURE STOCKÉE
-- ------------------------------------------------------------

CREATE OR ALTER PROCEDURE sp_GetEdgeAndRelated
    @key     INT,
    @top     INT = 100        -- Nombre max de lignes retournées (100 par défaut)
AS
BEGIN
    SET NOCOUNT ON;
    SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; -- Lecture sans verrous

    -- Variables locales pour stocker data1 / data2 de l'edge source
    DECLARE @d1 NVARCHAR(255);
    DECLARE @d2 NVARCHAR(255);

    -- --------------------------------------------------------
    -- Étape 1 : Récupérer data1 et data2 de l'edge source
    -- --------------------------------------------------------
    SELECT TOP 1
        @d1 = data1,
        @d2 = data2
    FROM edge WITH (NOLOCK)
    WHERE [key] = @key;

    IF @@ROWCOUNT = 0
    BEGIN
        RAISERROR('Edge introuvable pour la clé : %d', 16, 1, @key);
        RETURN;
    END

    -- --------------------------------------------------------
    -- Étape 2 : Retourner l'edge source (result set 1)
    -- --------------------------------------------------------
    SELECT
        [key],
        data1,
        data2,
        trans,
        output,
        code_ope
    FROM edge WITH (NOLOCK)
    WHERE [key] = @key;

    -- --------------------------------------------------------
    -- Étape 3 : Retourner les edges liés (result set 2)
    --   - TOP (@top)        : limite les lignes affichées
    --   - COUNT(*) OVER ()  : total global (avant la coupure TOP)
    --   - UNION ALL         : un seek par branche (plus performant que OR)
    --   - ORDER BY [key]    : résultat déterministe
    -- --------------------------------------------------------
    SELECT TOP (@top)
        [key],
        data1,
        data2,
        trans,
        output,
        code_ope,
        COUNT(*) OVER () AS TotalLignes
    FROM (
        -- Edges dont data1 = @d1
        SELECT [key], data1, data2, trans, output, code_ope
        FROM edge WITH (NOLOCK)
        WHERE data1 = @d1
          AND [key] <> @key

        UNION ALL

        -- Edges dont data1 = @d2 (si @d2 différent de @d1)
        SELECT [key], data1, data2, trans, output, code_ope
        FROM edge WITH (NOLOCK)
        WHERE data1 = @d2
          AND [key] <> @key
          AND @d2 <> @d1

        UNION ALL

        -- Edges dont data2 = @d1
        SELECT [key], data1, data2, trans, output, code_ope
        FROM edge WITH (NOLOCK)
        WHERE data2 = @d1
          AND [key] <> @key

        UNION ALL

        -- Edges dont data2 = @d2 (si @d2 différent de @d1)
        SELECT [key], data1, data2, trans, output, code_ope
        FROM edge WITH (NOLOCK)
        WHERE data2 = @d2
          AND [key] <> @key
          AND @d2 <> @d1

    ) AS edges_lies
    ORDER BY [key];

END;
GO

-- ------------------------------------------------------------
-- 3. EXEMPLES D'UTILISATION
-- ------------------------------------------------------------

-- 100 lignes par défaut
-- EXEC sp_GetEdgeAndRelated @key = 42;

-- 50 lignes
-- EXEC sp_GetEdgeAndRelated @key = 42, @top = 50;

-- 500 lignes
-- EXEC sp_GetEdgeAndRelated @key = 42, @top = 500;
