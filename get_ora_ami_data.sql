SELECT
  aoe_down_date,
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
  aoe_insvc_job_id,
  null as cause_code
FROM
  ami_outage_events @sdr
WHERE
  aoe_down_date >= add_months(trunc(sysdate), -12)
  AND aoe_down_date <= sysdate
  AND aoe_duration_minutes > 0
  AND aoe_duration_minutes <= 5
  AND aoe_is_meterremoval is null
  AND aoe_premise in ({})