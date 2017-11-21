import psycopg2
#add your postgres conenction information here
def connect():
	return psycopg2.connect("host='' dbname='' user='' password=''")
