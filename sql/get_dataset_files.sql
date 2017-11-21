Select contenttype, filesize from dvobject 
join datafile on datafile.id = dvobject.id
where owner_id = {id}
