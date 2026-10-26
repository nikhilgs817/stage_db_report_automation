-- Connectivity and formatting smoke test for Db2 on z/OS.
-- Replace with reviewed SELECT queries in the separate report repository.
SELECT CURRENT DATE AS REPORT_DATE,
       CURRENT SERVER AS DB2_LOCATION,
       CURRENT USER AS SERVICE_ACCOUNT
FROM SYSIBM.SYSDUMMY1;
