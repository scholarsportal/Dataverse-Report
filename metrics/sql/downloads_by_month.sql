WITH RECURSIVE tree(child, root) AS (
   SELECT c.id, c.owner_id
   FROM dvobject c
   LEFT JOIN dvobject p ON c.owner_id = p.id
   WHERE p.id = {object_id}
   UNION
   SELECT id,root
   FROM tree
   INNER JOIN dvobject on tree.child = dvobject.owner_id
)

SELECT *
FROM  (
   SELECT day::date
   FROM   generate_series(timestamp '{start_date}'
                        , timestamp '{end_date}'
                        , interval  '1 month') day
   ) d
LEFT   JOIN (
select date_trunc('month', responsetime)::date AS day,count(downloadtype)
from guestbookresponse, filedownload
where guestbookresponse.id = filedownload.guestbookresponse_id AND downloadtype like '%ownload%' and dataset_id IN (SELECT child FROM tree)
   GROUP  BY 1
   ) t USING (day)
ORDER  BY day;