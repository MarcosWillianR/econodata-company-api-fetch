from dotenv import load_dotenv
from  rpa_utils import WebScraperUtils
import requests
import os
import json

# Load environment variables from the .env file
load_dotenv()

class EconodataGetCompanies:
    def __init__(self):   
        self.utils = WebScraperUtils()

        self.token = ""
        self.base_url = "https://api.gke.econodata.com.br"
        self.default_headers = { 'Content-Type': 'application/json' }

        self.current_fetched_companies_count = 0
        self.total_companies_found = None

    def __load_request_payload_json(self, url):
        if url == "searchCompanies":
            with open('constants/search_company.json', 'r', encoding='utf-8') as file:
                payload = json.load(file)
            return payload
        elif url == "login":
            return { 
                "username": os.getenv("USER_NAME"), 
                "password": os.getenv("USER_PASSWORD")
            }

    def __login(self):
        try:
            r = requests.post(f'{self.base_url}/ecdt-user/api/user/login',
                              data = json.dumps(self.__load_request_payload_json("login")),
                              headers= self.default_headers)

            r.raise_for_status()
            data_json = r.json()

            response_token = data_json["access_token"]
            self.token = response_token
            self.default_headers["Cookie"] = f'ecdt_token=Bearer {response_token}'

            return True
        except requests.exceptions.RequestException as e:
            print(f"[__login]: An error ocurred: {e}")
            return False

    def __format_companies(self, companies):
        companies_formatted = []

        for company in companies:
            company_data = company["_source"]

            if "vlr_capital_social" not in company_data:
                continue

            if self.utils.cnpj_exists(company_data["num_cnpj"]):
                continue

            company_formatted = {
                "capital_social": company_data["vlr_capital_social"],
                "cnpj": company_data["num_cnpj"],
                "razao_social": company_data["razao_social"],
                "telefones": "",
                "pessoas": "",
            }

            if "ds_telefone" in company_data:
                phones = []
                try:
                    for phone in company_data["ds_telefone"]:
                        phone_formatted = f"{phone["ds_telefone"]} ({phone["tp_telefone"]})"
                        phones.append(phone_formatted)
                    company_formatted["telefones"] = "\n".join(phones)
                except:
                    pass
            
            if "cargos" in company_data:
                persons = []
                try:
                    for pessoa in company_data["cargos"]:
                        person_formatted = f"{pessoa["nome"]} ({pessoa["cargos"][0]})"
                        persons.append(person_formatted)
                    company_formatted["pessoas"] = "\n".join(persons)
                except:
                    pass
            
            companies_formatted.append(company_formatted)
        
        return companies_formatted

    def __save_companies_to_xlsx(self, companies):
        try:
            companies_formatted = self.__format_companies(companies)
            self.utils.xlsx_update(companies_formatted)
            self.current_fetched_companies_count += len(companies_formatted)
        except Exception as e:
            print(f"Error when save: {e}")

    def __get_companies(self, offset):
        try:
            payload = self.__load_request_payload_json("searchCompanies")
            payload["offsetRequest"] = offset

            r = requests.post(f'{self.base_url}/ecdt-busca/searchCompanies',
                             data = json.dumps(payload),
                             headers= self.default_headers)
            
            r.raise_for_status()
            data_json = r.json()

            self.total_companies_found = data_json["aggregations"]["totalEmpresas"]["value"]

            companies = data_json['hits']['hits']

            if len(companies) > 0:
                self.__save_companies_to_xlsx(companies)

            return True
        except requests.exceptions.RequestException as e:
            print(f"[__get_companies]: An error ocurred: {e}")
            return False

    def startX(self):
        success_login = self.__login()
        if success_login is False:
            return
        
        offset = 10
        items_processed = 10

        self.__get_companies(offset=0)

        while True:
            self.__get_companies(offset)
            if self.total_companies_found is None or self.current_fetched_companies_count >= self.total_companies_found:
                break
            
            items_processed += 10
            print(f'\rItens processados: {items_processed}', end='', flush=True)
            offset += 10
        print()
        

econodata = EconodataGetCompanies()
econodata.startX()