SELECT
  DISTINCT aoe_down_date,
  aoe_up_date,
  aoe_duration_minutes,
  aoe_circuit,
  last_value(aoe_circuit) OVER (
    PARTITION BY aoe_premise
    ORDER BY
      aoe_creation_date ROWS BETWEEN UNBOUNDED PRECEDING
      AND UNBOUNDED FOLLOWING
  ) as latest_circuit,
  aoe_premise,
  event_number as aoe_insvc_job_id,
  cause_code
FROM
  edw.ami_summary_v3
WHERE
  mtrrem = 0
  AND excludeconsec = 0
  AND AOE_DOWN_DATE between add_months(date_trunc('DAY', now()), -12)
  and now()
  AND audit_code not in ('Deleted', 'Blue Sky Cap')
  AND numcombined < 4
  AND aoe_circuit is not null
  AND aoe_premise in ({})