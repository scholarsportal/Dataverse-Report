Select COUNT(*) as total from guestbookresponse 
where downloadtype like '%ownload%' and dataset_id = {object_id}