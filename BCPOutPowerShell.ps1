$outTxt='c:\temp\output.txt'

$serverName='10.0.70.104'

$dbName='hudson_wl'

$queryStr="

DECLARE @as_of_dt DATE;
SET @as_of_dt = EOMONTH(GETDATE(), -1)

SELECT 'as_of_dt','loan_id','volt_bpo_asis_value','volt_bpo_dt','volt_security_name'
UNION ALL
SELECT  as_of_dt = cast(@as_of_dt as varchar(30)),
        cast(loan_id as varchar(30)),
        CAST(volt_bpo_asis_value as varchar(30)),
        cast(volt_bpo_dt as varchar(30)),
        cast(volt_security_name as varchar(30))
FROM    ( SELECT    e_prop.loan_id ,
                    volt_bpo_asis_value = ams.bpo_value ,
                    volt_bpo_dt = ams.value_dt ,
                    volt_security_name = ams.security_name ,
                    ROW_NUMBER() OVER ( PARTITION BY ams.loan_num_srvcr ORDER BY report_dt DESC ) AS rank
          FROM      loan.eq_ods_property e_prop
                    INNER JOIN dbo.AM_Securitization ams ON e_prop.loan_id = ams.loan_num_srvcr
          WHERE     e_prop.as_of_dt = @as_of_dt
        ) AS tbl
WHERE rank = 1;"

bcp $queryStr queryout $outTxt -S $serverName /d $dbName /c /t "|" /T