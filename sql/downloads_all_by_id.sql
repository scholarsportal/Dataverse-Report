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
select count(downloadtype)
from guestbookresponse
join filedownload on filedownload.guestbookresponse_id = guestbookresponse.id
join dvobject on dvobject.id = guestbookresponse.datafile_id
where dataset_id IN (SELECT child FROM tree)
and dvobject.publicationdate is not null;

