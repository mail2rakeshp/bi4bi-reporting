from tableau_api_lib import TableauServerConnection
tableau_server_config = {
        'tableau_prod': {
                'server': 'https://10ax.online.tableau.com',
                'api_version': '3.15',
                'username': 'tableaucoe2022@gmail.com',
                'password': 'Tableaucoe2022!',
                'site_name': 'Tableaucoe',
                'site_url': 'tableaucoedev394456'
        }
}
conn = TableauServerConnection(tableau_server_config)
conn.sign_in()
graphql_query = """
query getworkbooks {
  workbooks{id
  name}
  columnFields {
    id
    name
  }
}
"""
response = conn.metadata_graphql_query(query=graphql_query)
print(response.json())
conn.sign_out()