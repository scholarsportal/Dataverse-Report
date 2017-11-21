with recursive parent_table as (
  select  *
  from dvobject
  where id = {object_id}
  union
  select dvobject.*
  from dvobject 
  join parent_table on parent_table.owner_id = dvobject.id
)
select
  parent_table.id--, name
from parent_table 
WHERE  parent_table.id <> {object_id}
--left join  dataverse on parent_table.id = dataverse.id