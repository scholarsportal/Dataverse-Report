Select COUNT(*) as total from guestbookresponse, filedownload
where guestbookresponse.id = filedownload.guestbookresponse_id AND downloadtype like '%ownload%' and dataset_id = {object_id}