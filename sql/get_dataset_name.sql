Select value,identifier, protocol from datasetfield 
join datasetfieldvalue on datasetfieldvalue.datasetfield_id=datasetfield.id
join datasetversion on datasetfield.datasetversion_id = datasetversion.id
join  dataset on dataset.id= dataset_id
where datasetversion.id in (select max(datasetversion.id) as max from datasetversion group by datasetversion.dataset_id) and "datasetfieldtype_id" = 1 and dataset_id= {id}