# DynamicsPy
Realiza consultas utilizando python no CRM Dynamics 365, utilizando sua API web.
<br>
Se quiser tirar alguma dúvida pode me chamar no Linkedin https://www.linkedin.com/in/tadeu-mansi/

## Pré Requisitos
Instalar as dependências utilizadas no projeto
```python
pip install requests
pip install xml.etree.ElementTree
pip install urllib.parse
```
Ir no Active Directory no Azure e criar um aplicativo com as permissões de leitura de dados do Dynamics 365 (é necessário o concentimento do administrador do Azure), apos criar
o aplicativo obter o secret e o client id para uso de parâmetro na classe D365
## Como utilizar
Após instanciar a classe D365, você pode optar pelos métodos get_rows() e fetch_xml(), sendo respectivamente, get_rows() utilizado para realizar consultas utilziando Odata e
fetch_xml() utilizado para realizar consultas utilizando FetchXML, além de poder realizar uma Localização Avançada no Dynamics e exportar o seu FetchXML asssim o
utilizando no script python.

```python
from d365 import D365

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

```
