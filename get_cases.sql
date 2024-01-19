SELECT
    case_id,
    created_on,
    premise,
    frequent
FROM
    (
        SELECT
            case_id,
            created_on,
            premise,
            case
                when 
                    l2_category in ('FREQUENT OUTAGE', 'RESTORATION DURATION', 'STORM', 'OUTAGE ESTIMATES', 'RELIABILITY CREDITS')
                then 1
                else 0
            end as frequent,
            row_number() over (order by to_date(created_on, 'MM/DD/YYYY') DESC) rn
        FROM
            mt_all_cases_v2
        WHERE
            premise is not null
            AND l1_category LIKE 'DISTRI%'
            AND status = 2
            AND (case_category in ('MPSC', 'ICHP') or l2_category in ('FREQUENT OUTAGE', 'RESTORATION DURATION', 'STORM', 'OUTAGE ESTIMATES', 'RELIABILITY CREDITS'))
    )
WHERE
    rn <= 1000