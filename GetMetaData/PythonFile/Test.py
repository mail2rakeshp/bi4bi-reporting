import pandas as pd
import urllib
from sqlalchemy import create_engine
from pandas import json_normalize
from tableau_api_lib import TableauServerConnection
from tableau_api_lib.utils import flatten_dict_column, flatten_dict_list_column
tableau_server_config = {
         'tableau_prod': {
                 'server': '1',
                 'api_version': '1',
                 'username': '1',
                 'password': '1',
                 'site_name': '1',
                 'site_url':  '1'
         }
}
conn = TableauServerConnection(tableau_server_config)
conn.sign_in()
query_workbooks = """
{
  workbooks {
    workbook_name: name
    workbook_id: luid
    workbook_project: projectName
    views {
      view_type: __typename
      view_name: name
      view_id: luid
  }
  upstreamTables {
    upstr_table_name: name
    upstr_table_id: luid
    upstreamDatabases {
      upstr_db_name: name
      upstr_db_type: connectionType
      upstr_db_id: luid
      upstr_db_isEmbedded: isEmbedded
    }
  }
  upstreamDatasources {
    upstr_ds_name: name
    upstr_ds_id: luid
    upstr_ds_project: projectName
  }
  embeddedDatasources {
    emb_ds_name: name
  }
  upstreamFlows {
    flow_name: name
    flow_id: luid
    flow_project: projectName
  }
 }
}
"""
query_databases = """
{
  databaseServers {
    database_hostname: hostName
        database_port: port
    database_id: luid
  }
}
"""
wb_query_results = conn.metadata_graphql_query(query_workbooks)
db_query_results = conn.metadata_graphql_query(query_databases)
db_query_results_json = db_query_results.json()['data']['databaseServers']
wb_query_results_json = wb_query_results.json()['data']['workbooks']
wb_df = json_normalize(wb_query_results.json()['data']['workbooks'])
wb_df.drop(columns=['views', 'upstreamTables', 'upstreamDatasources', 'embeddedDatasources', 'upstreamFlows'], inplace=True)
wb_df=pd.DataFrame(wb_df)
wb_views_df = json_normalize(data=wb_query_results_json, record_path='views', meta='workbook_id')
wb_views_df=pd.DataFrame(wb_views_df)
wb_tables_df = json_normalize(data=wb_query_results_json, record_path='upstreamTables', meta='workbook_id')
wb_tables_df=pd.DataFrame(wb_tables_df)
wb_tables_dbs_df = flatten_dict_list_column(df=wb_tables_df, col_name='upstreamDatabases')
wb_tables_dbs_df=pd.DataFrame(wb_tables_dbs_df)
db_df = pd.DataFrame(db_query_results_json)
wb_uds_df = json_normalize(data=wb_query_results_json, record_path='upstreamDatasources', meta='workbook_id')
wb_uds_df=pd.DataFrame(wb_uds_df)
wb_eds_df = json_normalize(data=wb_query_results_json, record_path='embeddedDatasources', meta='workbook_id')
wb_eds_df=pd.DataFrame(wb_eds_df)
wb_flows_df = json_normalize(data=wb_query_results_json, record_path='upstreamFlows', meta='workbook_id')
wb_flows_df=pd.DataFrame(wb_flows_df)
quoted = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};SERVER=IN3040866W1\SQLEXPRESS;DATABASE=Power BI COE;Trusted_Connection=yes;")
engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))
wb_df.to_sql('TableauWorkbooks', schema='dbo', con = engine)
wb_views_df.to_sql('TableauViews', schema='dbo', con = engine)
db_df.to_sql('TableauDatabase', schema='dbo', con = engine)
wb_eds_df.to_sql('TableauEmbeddedDatasource', schema='dbo', con = engine)
