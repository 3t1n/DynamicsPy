import requests
import xml.etree.ElementTree as ET
import urllib.parse
import time

__author__ = "Tadeu Mansi"


class D365:
    """
    A class used to connect to the Dynamics 365 CRM.
    This class use Odata API

    A classe é usada para se conectar ao Dynamics 365 CRM
    Essa classe usa o Odata API

    Attributes
    ----------
    __client_id : str
        Is the client_id, the code of your azure application
        O client_id usado no seu aplicativo do azure
    __client_secret : str
        Is the client_secret, the token generate in azure application
        O client_secret, que você gera no seu aplicativo do azure
    __username : str
        Is the email used to connect in Dynamics 365 CRM
        O email usado para se conectar com o Dynamics 365 CRM
    __password : str
        Is the password used to connect in Dynamics 365 CRM
        A senha usada para se conectar com o Dynamics 365 CRM
    __crm_org : str
        Your Dynamics 365 URL, example: https://test.crm2.dynamics.com
        Sua URL do Dynamics 365 CRM, exemplo: https://test.crm2.dynamics.com
    __api_version : str
        Is your version of Web API
        Sua versão da Web API
    __crm_url : str
        Is complete url with API version
        Url Completa com a versão da API
    __header : str
        Is the header of requests used in code
        O cabeçalho usado para fazer as requisições no código

    Methods
    -------

    get_rows(self,query)
        Returns a list of Odata query lines
        Retorna uma lista com as linhas da consulta Odata

    fetch_xml(self, fetchxml)
        Returns a list of fetchxml query lines
        Retorna uma lista com as linhas da consulta fetchxml

    __request_crm(self, method, url, header, data=None, **kwargs)
        Send requests for Dynamics 365 CRM
        Envia requisições para o Dynamics 365 CRM

    __get_token(self)
        This method get the response request and returns token data
        Esse método pega a resposta da requisição e retorna os dados do token

    __parse_response(self, response)
        This method get the response request and returns json data or raise exceptions
        Esse método pega a resposta do request e retorna os dados em JSON
    """

    def __init__(self, username, password, client_id, crm_org, client_secret):
        self.__start_time = time.time()
        self.__client_id = client_id
        self.__client_secret = client_secret
        self.__username = username
        self.__password = password
        self.__crm_org = crm_org
        self.__api_version = "/api/data/v9.1/"
        self.__crm_url = crm_org + self.__api_version
        self.__header = {
            'Authorization': '',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json',
            'Content-Type': 'application/json; charset=utf-8'
        }

    def get_rows(self, query: str) -> list:
        """
        Returns a list of Odata query lines
        Retorna uma lista com as linhas da consulta Odata

        Parameters
        ----------
        query : str
            The query Odata passed in web API, example: contacts?$select=fullname
            A query Odata passada na web API, exemplo: contacts?$select=fullname

        """

        self.__header['Prefer'] = "odata.maxpagesize=5000"
        self.__header['Authorization'] = self.__get_token()
        header = self.__header
        if self.__api_version not in query:
            url = self.__crm_url + query
        else:
            url = query
        response = self.__request_crm('get', url, header)
        print("Page 1")
        if '@odata.nextLink' in response.keys():
            all_records = response
            url = response['@odata.nextLink']
            page = 0
            while True:
                page += 1
                print("Page " + str(page))
                response = self.__request_crm('get', url, header)
                entries = len(response['value'])
                count = 0
                while count < entries:
                    all_records['value'].append(response['value'][count])
                    count += 1
                if '@odata.nextLink' in response.keys():
                    url = response['@odata.nextLink']
                else:
                    break
            print("Lines " + str(len(all_records['value']) - 1))
            print("--- %s seconds ---" % (time.time() - self.__start_time))
            print("--- %s minutes ---" % ((time.time() - self.__start_time)/60))
            return all_records['value']
        else:
            if 'value' in response.keys():
                return response['value']
            else:
                return response

    def fetch_xml(self, fetchxml: str) -> list:
        """
        Returns a list of fetchxml query lines
        Retorna uma lista com as linhas da consulta fetchxml

        Parameters
        ----------
        fetchxml : str
            The query fetchxml passed in web API
            A query fetchxml passada na web API

        """
        self.__header['Prefer'] = "odata.include-annotations=*"
        self.__header['Authorization'] = self.__get_token()

        header = self.__header
        root = ET.fromstring(fetchxml)
        try:
            entidade = root[0].attrib['name']
        except KeyError:
            raise Exception("Could not get the name of entity")
       
        fetchxml = ET.tostring(root, encoding='utf8', method='xml')
        url_body = self.__crm_url + entidade + "s" + "?fetchXml="
        url = url_body + urllib.parse.quote(fetchxml)
        response_fisrt = self.__request_crm('get', url, header)
        print("Page 1")
        if '@Microsoft.Dynamics.CRM.fetchxmlpagingcookie' in response_fisrt.keys():
            xml_cookie = ET.fromstring(response_fisrt['@Microsoft.Dynamics.CRM.fetchxmlpagingcookie'])
            all_records = response_fisrt
            page = 2
            test = xml_cookie.attrib['pagingcookie']
            data = urllib.parse.unquote(urllib.parse.unquote(test))
            data.replace("&", '&amp;').replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;")
            root.set("paging-cookie", data)
            root.set("count", "5000")
            root.set("page", str(page))
            fetchxml = ET.tostring(root, encoding='utf8', method='xml')
            url = url_body + urllib.parse.quote(fetchxml)
            while True:
                print("Page " + str(page))
                page += 1
                response = self.__request_crm('get', url, header)
                entries = len(response['value'])
                count = 0
                while count < entries:
                    all_records['value'].append(response['value'][count])
                    count += 1
                if '@Microsoft.Dynamics.CRM.fetchxmlpagingcookie' in response.keys():
                    test = xml_cookie.attrib['pagingcookie']
                    data = urllib.parse.unquote(urllib.parse.unquote(test))
                    data.replace("&", '&amp;').replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;")
                    root.set("paging-cookie", data)
                    root.set("page", str(page))
                    fetchxml = ET.tostring(root, encoding='utf8', method='xml')
                    url = self.__crm_url + entidade + "s" + "?fetchXml=" + urllib.parse.quote(fetchxml)
                else:
                    break
            print("Lines " + str(len(all_records['value']) - 1))
            print("--- %s seconds ---" % (time.time() - self.__start_time))
            print("--- %s minutes ---" % ((time.time() - self.__start_time)/60))
            return all_records['value']
        else:
            if 'value' in response_fisrt.keys():
                return response_fisrt['value']
            else:
                return response_fisrt

    def __request_crm(self, method: str, url: str, header: dict, data=None, **kwargs):
        """
        Send requests for Dynamics 365 CRM
        Envia requisições para o Dynamics 365 CRM

        Parameters
        ----------
        method : str
            Method used in request
            Método usado na requisição
        url : str
            Url used in request
            Url usada na requisição
        header : dict
            Header of request
            Cabeçalho da requisição
        data : list
            Data in body request
            Dados no corpo da requisição

        """
        response = ""
        if method.lower() == "get":
            response = requests.get(url, headers=header, params=kwargs)
        elif method.lower() == "post":
            response = requests.post(url, headers=header, data=data)
        return self.__parse_response(response)

    def __get_token(self) -> str:
        """
        This method get the response request and returns token data
        Esse método pega a resposta da requisição e retorna os dados do token
        """

        tokenpost = {
            'client_id': self.__client_id,
            'resource': self.__crm_org,
            'username': self.__username,
            'password': self.__password,
            'client_secret': self.__client_secret,
            'grant_type': 'password'
        }
        response = self.__parse_response(
            requests.post("https://login.microsoftonline.com/common/oauth2/token", data=tokenpost))
        try:
            return response['access_token']
        except KeyError:
            raise Exception("Could not get access token")

    @staticmethod
    def __parse_response(response) -> list:
        """
        This method get the response request and returns json data or raise exceptions
        Esse método pega a resposta da requisição e retorna os dados em JSON

        Parameters
        ----------
        response : str
            The request response
            A resposta da requisição
        """
        if response.status_code == 204 or response.status_code == 201:
            return True
        elif response.status_code == 400:
            raise Exception(
                "The URL {0} retrieved an {1} error. "
                "Please check your request body and try again.\nRaw message: {2}".format(
                    response.url, response.status_code, response.text))
        elif response.status_code == 401:
            raise Exception(
                "The URL {0} retrieved and {1} error. Please check your credentials, make sure you have permission to "
                "perform this action and try again.".format(
                    response.url, response.status_code))
        elif response.status_code == 403:
            raise Exception(
                "The URL {0} retrieved and {1} error. Please check your credentials, make sure you have permission to "
                "perform this action and try again.".format(
                    response.url, response.status_code))
        elif response.status_code == 404:
            raise Exception(
                "The URL {0} retrieved an {1} error. Please check the URL and try again.\nRaw message: {2}".format(
                    response.url, response.status_code, response.text))
        elif response.status_code == 412:
            raise Exception(
                "The URL {0} retrieved an {1} error. Please check the URL and try again.\nRaw message: {2}".format(
                    response.url, response.status_code, response.text))
        elif response.status_code == 413:
            raise Exception(
                "The URL {0} retrieved an {1} error. Please check the URL and try again.\nRaw message: {2}".format(
                    response.url, response.status_code, response.text))
        elif response.status_code == 500:
            raise Exception(
                "The URL {0} retrieved an {1} error. Please check the URL and try again.\nRaw message: {2}".format(
                    response.url, response.status_code, response.text))
        elif response.status_code == 501:
            raise Exception(
                "The URL {0} retrieved an {1} error. Please check the URL and try again.\nRaw message: {2}".format(
                    response.url, response.status_code, response.text))
        elif response.status_code == 503:
            raise Exception(
                "The URL {0} retrieved an {1} error. Please check the URL and try again.\nRaw message: {2}".format(
                    response.url, response.status_code, response.text))
        return response.json()
