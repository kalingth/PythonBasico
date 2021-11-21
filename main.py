import numpy as np
import requests as req
from bs4 import BeautifulSoup
import pandas as pd
import matplotlib.pyplot as plt
from ast import literal_eval as str2l
from re import search
from datetime import datetime


class Covidata:
    @staticmethod
    def get_data():
        response = req.get("https://www.worldometers.info/coronavirus/")
        soup = BeautifulSoup(response.text, 'html.parser')
        table = soup.find(id="main_table_countries_today").tbody
        headers = ["Position", "Country", "TotalCases", "NewCases", "TotalDeaths", "NewDeaths", "TotalRecovered",
                   "NewRecovered", "ActiveCases", "SeriousCritical", "CasesMillion", "DeathsMillion",
                   "TotalTests", "TestsMillion", "Population", "Region"]
        buffer = [[data.text.strip().replace(",", "")
                   for data in list(row.find_all("td"))[:16]]
                  for row in table.find_all("tr")]
        dataset = pd.DataFrame(buffer, columns=headers)
        dataset = dataset.replace("", np.NAN).replace("N/A", np.NAN)
        dataset = dataset[dataset["Position"].notnull()]
        for key in headers:
            if key != "Country" and key != "Region":
                dataset[key] = pd.to_numeric(dataset[key])
        return dataset

    @staticmethod
    def prepareJson(script):
        dataRegex = r"\b(data:).+\]"
        categoryRegex = r"\b(categories:).+\]"
        catRemove = "categories:"
        dataRemove = "data:"
        script = script.replace("}]", "")
        recCat = search(categoryRegex, script)[0].replace(catRemove, "").replace("null", "None")
        recData = search(dataRegex, script)[0].replace(dataRemove, "").replace("null", "None")

        extracted = {
            "dates": list(map(lambda x: datetime.strptime(x, "%b %d, %Y"), str2l(recCat))),
            "numbers": str2l(recData)
        }
        return extracted

    def show_graph(self, country=None, type="TotalCases"):
        if country:
            toPlot = self.countries[country][type]
            toPlot.plot(kind="line", y="numbers", x="dates")
        else:
            toPlot = self.globalData.head(10)
            toPlot.plot(kind="bar", x="Country", y=type)
        plt.show()

    def export(self, name):

        with pd.ExcelWriter(name, engine='xlsxwriter') as writer:
            self.globalData.to_excel(writer, sheet_name='Global', index=False)

            workbook = writer.book

            to_graphs = {"TotalCases": "C", "NewCases": "D", "TotalDeaths": "E", "TotalRecovered": "G"}

            for graph in to_graphs.keys():
                worksheet = workbook.add_worksheet(graph)
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': f'=Global!${to_graphs[graph]}$1',
                    'values': f'=Global!${to_graphs[graph]}$2:${to_graphs[graph]}$20',
                    'categories': '=Global!$B$2:$B$20'
                })
                worksheet.insert_chart('D2', chart, {'x_scale': 3, 'y_scale': 2})

        print(f"Data saved as {name}")

    def export_country(self, country):
        with pd.ExcelWriter(f"{country}.xlsx", engine='xlsxwriter') as writer:
            for key in self.countries[country].keys():
                data = self.countries[country][key]
                data.to_excel(writer, sheet_name=key, index=False)
                workbook = writer.book

                worksheet = workbook.get_worksheet_by_name(key)
                chart = workbook.add_chart({'type': 'line'})
                chart.set_x_axis({'num_format': 'DD/MM/YY'})
                chart.set_y_axis({'num_format': "##0.0 E+0"})
                chart.add_series({
                    'name': key,
                    'values': f'={key}!$B$2:$B${len(data)}',
                    'categories': f'={key}!$A$2:$A${len(data)}'
                })
                worksheet.insert_chart('D2', chart, {'x_scale': 2, 'y_scale': 1.5})

    def get_country_data(self):
        for country in self.countries.keys():
            response = req.get(f"https://www.worldometers.info/coronavirus/country/{country}/")
            soup = BeautifulSoup(response.text, 'html.parser')

            wanted = {
                "TotalCases": "#graph-active-cases-total",
                "TotalDeaths": ".tabbable-panel-deaths",
                "DailyDeath": "#graph-deaths-daily",
                "CuredDaily": "#cases-cured-daily"
            }

            for wKey in wanted.keys():
                elem = soup.select(wanted[wKey])[0].parent
                script = elem.script.text
                self.countries[country][wKey] = pd.DataFrame(self.prepareJson(script))
                if wKey == "TotalCases":
                    self.countries[country][wKey]["daily data"] = self.countries[country][wKey]["numbers"]
                    self.countries[country][wKey]["numbers"] = self.countries[country][wKey]["numbers"].cumsum()

    def __init__(self, *args):
        if args:
            self.countries = {country: {} for country in args}
            self.get_country_data()
        self.globalData = self.get_data()


if __name__ == "__main__":
    cov2an = Covidata("brazil", "italy", "germany", "japan")
    cov2an.export_country("brazil")
    cov2an.export_country("italy")
    cov2an.export("world.xlsx")
    cov2an.show_graph()
    cov2an.show_graph("italy")
