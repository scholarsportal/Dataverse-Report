Select dvobject.id as root_id, dtype, publicationdate from dvobject
where dvobject.owner_id = {dataverse_id} and
dtype = 'Dataverse' and
publicationdate IS NOT NULL;