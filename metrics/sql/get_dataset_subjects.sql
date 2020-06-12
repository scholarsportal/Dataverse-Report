select distinct strvalue,datasetfield_controlledvocabularyvalue.controlledvocabularyvalues_id, 
	datasetversion.id, unf, releasetime, versionstate, dataset_id
from datasetfield_controlledvocabularyvalue 
join controlledvocabularyvalue
	on controlledvocabularyvalue.id = datasetfield_controlledvocabularyvalue.controlledvocabularyvalues_id
join datasetfield
	on datasetfield.id = datasetfield_controlledvocabularyvalue.datasetfield_id
join datasetversion
	on datasetversion.id = datasetfield.datasetversion_id
where controlledvocabularyvalue.datasetfieldtype_id = 19 and dataset_id = {id}
order by datasetfield_controlledvocabularyvalue.controlledvocabularyvalues_id asc; 