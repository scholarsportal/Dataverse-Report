SELECT *
FROM  (
   SELECT day::date
   FROM   generate_series(timestamp '{start_date}'
                        , timestamp '{end_date}'
                        , interval  '1 month') day
   ) d
LEFT   JOIN (
select date_trunc('month', responsetime)::date AS day,count(downloadtype) 
from guestbookresponse 
where downloadtype like '%ownload%' and dataset_id ={object_id} 
   GROUP  BY 1
   ) t USING (day)
ORDER  BY day;