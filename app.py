import pymysql
import xlrd
import logging
import sys

rds_host = '******************************' 
name = '********' 
password = '***********' 
db_name = '*********' 

logger = logging.getLogger()
logger.setLevel(logging.INFO)
    
try:
    conn = pymysql.connect(rds_host, user=name, passwd=password, db=db_name, connect_timeout=5)
except pymysql.MySQLError as e:
    logger.error("ERROR: Unexpected error: Could not connect to MySQL instance.")
    logger.error(e)
    sys.exit()

def handler(event, context): 
    filePath="040_4a_R5.xls"
    openFile=xlrd.open_workbook(filePath)
    sheet=openFile.sheet_by_name("MINFO")

    item_count = 0

    with conn.cursor() as cur:
        #Ya tenemos creada esta base de datos
        #cur.execute("create table CREATE TABLE Banca_Multiple (Estado varchar (20), Municipio Varchar (20), Sucursal int);")
        
        for i in range(12, sheet.nrows):
            cur.execute('insert into Banca_Multiple (Estado, Municipio, Sucursal) values(' 
            + sheet.cell_value(i,1) + ", " + sheet.cell_value(i,2) + ", " + sheet.cell_value(i,3) + ")")
            conn.commit()
            print(sheet.cell_value(i,1) + ", " + sheet.cell_value(i,2) + ", " + sheet.cell_value(i,3));

        cur.execute("select * from Employee")
        for row in cur:
            item_count += 1
            logger.info(row)
            #print(row)
    conn.commit()

    return "Added %d items from RDS MySQL table" %(item_count)
