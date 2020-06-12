SELECT *
FROM  (
   SELECT day::date
   FROM   generate_series(timestamp '{start_date}'
                        , timestamp '{end_date}'
                        , interval  '1 month') day
   ) d
LEFT   JOIN (
select date_trunc('month', createdtime)::date AS day,count(*), affiliation 
from authenticateduser 
   GROUP  BY 1,affiliation
   ) t USING (day)
ORDER  BY day;