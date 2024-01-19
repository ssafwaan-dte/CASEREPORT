select
  c.premise_number,
  c.customer_name,
  c.phone,
  c.meter_number,
  c.service_address
from
  edw.customer_premise_primary_meter c
where
  c.premise_number in ({})