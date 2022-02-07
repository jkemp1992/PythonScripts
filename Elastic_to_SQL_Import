import logging
import pandas as pd
import sys
from elasticsearch import Elasticsearch
import elasticsearch.helpers
import json
import pyodbc


# Function that replaces unwanted characters in a string
def function_replace(mainString, toBeReplaces, newString):
    # Iterate over the strings to be replaced
    for elem in toBeReplaces:
        # Check if string is in the main string
        if elem in mainString:
            # Replace the string
            mainString = mainString.replace(elem, newString)
    return mainString


# Function that prepares strings for MS SQL use
def clean_string(my_string):
    # If Variable is empty, replace with NULL
    if str(my_string) == "nan" or str(my_string) == "":
        result_string = "NULL"
    else:
        # Replace quotes and commas
        result_string = function_replace(str(my_string), ["'", '"', ","], "")
        result_string = "'" + result_string + "'"
    return result_string


# References
# https://www.elastic.co/guide/en/elasticsearch/client/python-api/current/index.html
# https://elasticsearch-py.readthedocs.io/en/master/index.html
# https://elasticsearch-py.readthedocs.io/en/latest/api.html#elasticsearch


def ElasticSearchAPICall(index_name, query_body, query_fields):
    fmt = "%(asctime)s %(levelname)s: %(message)s"
    logging.basicConfig(format=fmt, level=logging.CRITICAL)  # Can set to critical, debug, info, error, fatal
    logger = logging.getLogger(__name__)

    query_body["fields"] = query_fields
    query_body["_source"] = False

    try:
        logger.info("Connecting to Elasticsearch")
        es = Elasticsearch(
            [
                # Known server names/IPs here
            ],
            api_key=("super", "secret"),
            max_retries=5,
            retry_on_timeout=True,
            scheme="https",
            sniff_on_start=False,
            sniff_on_connection_fail=True,
            ssl_show_warn=False,
            timeout=5,
            use_ssl=True,
            verify_certs=False,
            http_compress=True,
        )
        logger.info(json.dumps(es.info(), indent=2))
    except Exception as ex:
        sys.exit(logger.error(ex))

    # results = elasticsearch.helpers.scan(es, index=index_name, query=query_body, size=10000)
    results = []
    # count = 0
    # sys.stdout.write(f"Total Records:       {count}")
    # sys.stdout.flush()
    for record in elasticsearch.helpers.scan(es, index=index_name, query=query_body, size=10000):
        results.append(record)
        # count += 1
        # if count % 10000 == 0:
        #     sys.stdout.write("\b" * len(str(count)))
        #     sys.stdout.write(f"{count}")
        #     sys.stdout.flush()

    # sys.stdout.write("\b" * len(str(count)))
    # sys.stdout.write(f"{count}")
    sys.stdout.write(f"Total Records: {len(results):,}\n")

    # Loop through results and add to list
    my_list = []
    for item in results:
        # Read the item as a dictionary
        my_dict = item["fields"]
        # Loop through dictionaries and replace [''] around each value
        for key, value in my_dict.items():
            my_dict[key] = function_replace(str(value), ["['", "']", "[", "]"], "")
        # Add dictionaries to a list
        my_list.append(my_dict)

    # Create DataFrame, remove duplicates
    df = pd.DataFrame(my_list, columns=query_fields)
    df.drop_duplicates(inplace=True)

    return df


sys.stdout.write("Pulling from Elastic: Example\n")

def getElasticData():
    index = "example.elastic.index*"
    query = { # Elastic query here
        "query": {
            "bool": {
                "must": [],
                "filter": [],
            }
        }
    }
    my_fields = ['Field1', 'Field2']
    data_table = ElasticSearchAPICall(index, query, my_fields)
    # data_table.to_csv("Example_Data.csv")

    # Connection to SQL
    SQL_Server_CN = pyodbc.connect('Driver={SQL Server};'
                                   'Server=MY-SQLSERVER;'
                                   'Database=MY-DB;'
                                   'UID=MY-USERNAME;'
                                   'PWD={MY-PASS};'
                                   'Trusted_Connection=no;'
                                   )

    cursor = SQL_Server_CN.cursor()

    # Truncate table for fresh import
    SQLString = "SET NOCOUNT ON; DELETE FROM MY_TABLE "
    cursor.execute(SQLString)
    cursor.commit()
    SQLString = "BEGIN TRANSACTION \r\n"
    count = 0

    sys.stdout.write("----Importing into SQL Database\n")
    sys.stdout.write(f"----Records Imported:        {count}")
    sys.stdout.flush()

    # Loop through data table
    # Clean each string to prepare for SQL import
    for x in range(0, len(data_table)):
        Field1 = clean_string(data_table.iloc[x]['Field1'])
        Field2 = clean_string(data_table.iloc[x]['Field2'])
        Field3 = clean_string(data_table.iloc[x]['Field3'])
        Field4 = clean_string(data_table.iloc[x]['Field4'])
        Field5 = clean_string(data_table.iloc[x]['Field5'])
        Field6 = clean_string(data_table.iloc[x]['Field6'])

        # Start the SQL batch import
        SQLString += f'''
                        EXEC [spi_SQL_Stored_Procedure]
                        @Field1 = {Field1}, 
                        @Field2 = {Field2}, 
                        @Field3 = {Field3}, 
                        @Field4 = {Field4}, 
                        @Field5 = {Field5}, 
                        @Field6 = {Field6} 

        '''
        count += 1
        # Import every 400 records by committing the transaction
        if count % 400 == 0:
            SQLString += "COMMIT TRANSACTION"
            try:
                cursor.execute(SQLString)
                cursor.commit()
                sys.stdout.write("\b" * len(str(count)))
                sys.stdout.write(f"{count}")
                sys.stdout.flush()
            # Catch and display any errors that arise
            except pyodbc.Error as err:
                print(f'Error: {err}')
                print(SQLString)
                sys.exit()
            # Start the batch import over
            SQLString = "BEGIN TRANSACTION \r\n"
            
    # Import the last batch
    if SQLString != "BEGIN TRANSACTION \r\n":
        SQLString += "COMMIT TRANSACTION"
        try:
            cursor.execute(SQLString)
            cursor.commit()
            sys.stdout.write("\b" * len(str(count)))
            sys.stdout.write(f"{count}")
            sys.stdout.flush()
        except pyodbc.Error as err:
            print(f'Error: {err}')
            print(SQLString)
            sys.exit()


    # Clean up and close connection
    cursor.close()
    del cursor
    SQL_Server_CN.close()


if __name__ == "__main__":
    getElasticData()
