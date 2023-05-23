from report.src.operators.xlsx_report_plugin import ExcelReportPlugin
import os
import json

base_path = os.sep.join(os.getcwd().split(os.sep)[:-3])
print(f'base path: {base_path}')

input_file = base_path + '/input_data/supermarket_sales.xlsx'
output_file = base_path + '/output_data/daily_gross_revenue_report.xlsx'

# Opening JSON file
# configs = open(base_path + '/configs/webhook.json')
# webhook_url = json.load(configs)['webhook_url']

automate = ExcelReportPlugin(
    input_file=input_file,
    output_file=output_file
)

if __name__ == "__main__":
    automate.main()
