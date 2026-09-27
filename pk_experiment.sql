-- =====================================================================
--  MySQL experiment: table WITH primary key vs WITHOUT primary key
--  Pure-SQL version (for MySQL Workbench or the mysql command line).
--  Timings: run "SHOW PROFILES;" after each block - the Duration column
--  is in seconds with microsecond precision.
-- =====================================================================

DROP DATABASE IF EXISTS pk_experiment_sql;
CREATE DATABASE pk_experiment_sql;
USE pk_experiment_sql;

SET SESSION cte_max_recursion_depth = 2000000;
SET SESSION profiling_history_size = 100;

-- ---------------------------------------------------------------------
-- STEP 1: two tables with the same columns
-- ---------------------------------------------------------------------
CREATE TABLE data_pk (
    id          INT            NOT NULL,
    name        VARCHAR(50)    NOT NULL,
    email       VARCHAR(100)   NOT NULL,
    age         INT            NOT NULL,
    city        VARCHAR(30)    NOT NULL,
    salary      DECIMAL(12,2)  NOT NULL,
    created_at  DATETIME       NOT NULL,
    PRIMARY KEY (id)                       -- <== the only difference
) ENGINE=InnoDB;

CREATE TABLE data_nopk (
    id          INT            NOT NULL,   -- normal column: no primary key, no index
    name        VARCHAR(50)    NOT NULL,
    email       VARCHAR(100)   NOT NULL,
    age         INT            NOT NULL,
    city        VARCHAR(30)    NOT NULL,
    salary      DECIMAL(12,2)  NOT NULL,
    created_at  DATETIME       NOT NULL
) ENGINE=InnoDB;

-- helper: builds a staging table of random rows with ids first_id .. last_id
-- (so both tables receive exactly the same random data)
DROP PROCEDURE IF EXISTS make_random_rows;
DELIMITER //
CREATE PROCEDURE make_random_rows(IN first_id INT, IN last_id INT)
BEGIN
    DROP TABLE IF EXISTS staging;
    CREATE TABLE staging ENGINE=InnoDB AS
    WITH RECURSIVE seq(n) AS (
        SELECT first_id UNION ALL SELECT n + 1 FROM seq WHERE n < last_id
    )
    SELECT n                                                     AS id,
           CONCAT('user_', SUBSTRING(MD5(RAND()), 1, 8))         AS name,
           CONCAT('user', n, '@example.com')                     AS email,
           18 + FLOOR(RAND() * 48)                               AS age,
           ELT(1 + FLOOR(RAND() * 12), 'Yogyakarta','Jakarta','Surabaya','Bandung','Medan','Semarang',
               'Makassar','Denpasar','Malang','Palembang','Solo','Bogor')  AS city,
           ROUND(3000000 + RAND() * 47000000, 2)                 AS salary,
           TIMESTAMP('2020-01-01') + INTERVAL FLOOR(RAND() * 189216000) SECOND AS created_at
    FROM seq;
END //
DELIMITER ;

-- ---------------------------------------------------------------------
-- STEP 2: insert 1,000,000 random rows into each table
-- ---------------------------------------------------------------------
CALL make_random_rows(1, 1000000);
SET profiling = 1;
INSERT INTO data_pk   SELECT * FROM staging;
INSERT INTO data_nopk SELECT * FROM staging;
SET profiling = 0;
SHOW PROFILES;                              -- <== note the two INSERT durations

-- ---------------------------------------------------------------------
-- STEP 3: select row 10, 1,000, 10,000, 100,000, 1,000,000
-- ---------------------------------------------------------------------
SET profiling = 1;
SELECT * FROM data_pk   WHERE id = 10;
SELECT * FROM data_nopk WHERE id = 10;
SELECT * FROM data_pk   WHERE id = 1000;
SELECT * FROM data_nopk WHERE id = 1000;
SELECT * FROM data_pk   WHERE id = 10000;
SELECT * FROM data_nopk WHERE id = 10000;
SELECT * FROM data_pk   WHERE id = 100000;
SELECT * FROM data_nopk WHERE id = 100000;
SELECT * FROM data_pk   WHERE id = 1000000;
SELECT * FROM data_nopk WHERE id = 1000000;
SET profiling = 0;
SHOW PROFILES;                              -- <== note the 10 SELECT durations

-- why: look at "type" (const = index lookup, ALL = full table scan) and "rows"
EXPLAIN SELECT * FROM data_pk   WHERE id = 1000000;
EXPLAIN SELECT * FROM data_nopk WHERE id = 1000000;

-- ---------------------------------------------------------------------
-- STEP 4 + 5: insert 20,000 / 40,000 / 60,000 / 100,000 more rows
-- ---------------------------------------------------------------------
CALL make_random_rows(1000001, 1020000);   -- 20,000 rows
SET profiling = 1;
INSERT INTO data_pk   SELECT * FROM staging;
INSERT INTO data_nopk SELECT * FROM staging;
SET profiling = 0;

CALL make_random_rows(1020001, 1060000);   -- 40,000 rows
SET profiling = 1;
INSERT INTO data_pk   SELECT * FROM staging;
INSERT INTO data_nopk SELECT * FROM staging;
SET profiling = 0;

CALL make_random_rows(1060001, 1120000);   -- 60,000 rows
SET profiling = 1;
INSERT INTO data_pk   SELECT * FROM staging;
INSERT INTO data_nopk SELECT * FROM staging;
SET profiling = 0;

CALL make_random_rows(1120001, 1220000);   -- 100,000 rows
SET profiling = 1;
INSERT INTO data_pk   SELECT * FROM staging;
INSERT INTO data_nopk SELECT * FROM staging;
SET profiling = 0;
SHOW PROFILES;                              -- <== note the 8 INSERT durations (PK, no-PK per batch)

SELECT COUNT(*) FROM data_pk;              -- 1,220,000
SELECT COUNT(*) FROM data_nopk;            -- 1,220,000
DROP TABLE staging;
