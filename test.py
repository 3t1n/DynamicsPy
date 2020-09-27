from d365 import D365
import pandas as pd

crm = D365("crm@email.com", "password", "azure_client_id",
          "https://yourcrm.crm2.dynamics.com", "azure_secret")
 
#geting results using odata
result_odata = crm.get_rows("accounts(E4F3B6D5-9F6A-427D-9592-31B6CD877E77)")

fetchxml = '''
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
 <entity name="account">
   <attribute name="name" />
   <attribute name="primarycontactid" />
   <attribute name="telephone1" />
   <attribute name="accountid" />
</entity>
</fetch>
'''

#geting results using fetchxml
result_fetch_xml = crm.fetch_xml(fetchxml)

#geting results and puting in csv file
df = pd.DataFrame(result_fetch_xml)
df.drop('@odata.etag',axis=1,inplace=True)
df.to_csv("data.csv", sep='\t', encoding='utf-8', index=False)
