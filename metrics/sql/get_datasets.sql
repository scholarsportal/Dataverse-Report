Select dvobject.id as id, dtype,publicationdate from dvobject 
where dtype = 'Dataset' and publicationdate IS NOT NULL
Order by id