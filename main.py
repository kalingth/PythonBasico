import numpy as np
import requests as req
from bs4 import BeautifulSoup
import pandas as pd
import matplotlib.pyplot as plt

class Covidata:
    @staticmethod
    def get_data():
        response = req.get("https://www.worldometers.info/coronavirus/")
        soup = BeautifulSoup(response.text, 'html.parser')
        table = soup.find(id="main_table_countries_today").tbody
        headers = ["Position", "Country", "TotalCases", "NewCases", "TotalDeaths", "NewDeaths", "TotalRecovered",
                   "NewRecovered", "ActiveCases", "SeriousCritical", "CasesMillion", "DeathsMillion",
                   "TotalTests", "TestsMillion", "Population", "Region"]
        buffer = [[data.text.strip().replace(",","")
                   for data in list(row.find_all("td"))[:16]]
                  for row in table.find_all("tr")]
        dataset = pd.DataFrame(buffer, columns=headers)
        dataset = dataset.replace("", np.NAN).replace("N/A", np.NAN)
        dataset = dataset[ dataset["Position"].notnull() ]
        for key in headers:
            if key != "Country" and key != "Region":
                dataset[key] = pd.to_numeric(dataset[key])
        return dataset

    def show_graph(self, type="TotalCases"):
        toPlot = self.dataset.head(10)
        toPlot.plot(kind="bar", x="Country", y=type)
        plt.show()

    def export(self, name):

        with pd.ExcelWriter(name, engine='xlsxwriter') as writer:
            self.dataset.to_excel(writer, sheet_name='Dados' , index=False)

            workbook = writer.book

            to_graphs = {"TotalCases":"C", "NewCases":"D", "TotalDeaths":"E", "TotalRecovered":"G"}

            for graph in to_graphs.keys():
                worksheet = workbook.add_worksheet(graph)
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': f'=Dados!${to_graphs[graph]}$1',
                    'values': f'=Dados!${to_graphs[graph]}$2:${to_graphs[graph]}$20',
                    'categories': '=Dados!$B$2:$B$20'
                })
                worksheet.insert_chart('D2', chart, {'x_scale': 3, 'y_scale': 2})


        print(f"Data saved as {name}")

    def __init__(self):
        self.dataset = self.get_data()


if __name__ == "__main__":
    cov2an = Covidata()
    cov2an.show_graph()
    cov2an.export("output.xlsx")