from collections import defaultdict
from odoo import models
import base64
import xlsxwriter
from io import BytesIO

class WarehouseAnalysisReport(models.AbstractModel):
    _name = 'report.warehouse_analysis_report.warehouse_analysis_template'
    _description = 'Warehouse Analysis Report'
    _inherit = "report.report_xlsx.abstract"

    def generate_xlsx_report(self, workbook, data, lines):
        start_date = data.get('start')
        end_date = data.get('end')
        reports = self.env['stock.report'].search([
            ('date_done', '>=', start_date),
            ('date_done', '<=', end_date)
        ])
        grouped_data = defaultdict(lambda: {
            'company_name': '',
            'partners': [],
            'total_delay': 0,
            'total_cycle_time': 0,
            'total_product_qty': 0,
            'total_count': 0
        })

        for report in reports:
            company_id = report.company_id.id if report.company_id else 'No Company'
            company_name = report.company_id.name if report.company_id else 'No Company'
            
            partner_name = report.partner_id.name if report.partner_id else 'No Partner'
            grouped_data[company_id]['company_name'] = company_name
            grouped_data[company_id]['partners'].append(partner_name)
            grouped_data[company_id]['total_delay'] += report.delay or 0
            grouped_data[company_id]['total_cycle_time'] += report.cycle_time or 0
            grouped_data[company_id]['total_product_qty'] += report.product_qty or 0
            grouped_data[company_id]['total_count'] += 1
        all_companies = self.env['res.company'].search([])
        for company in all_companies:
            if company.id not in grouped_data:
                grouped_data[company.id] = {
                    'company_name': company.name,
                    'partners': [],
                    'total_delay': 0,
                    'total_cycle_time': 0,
                    'total_product_qty': 0,
                    'total_count': 0
                }
        worksheet = workbook.add_worksheet('Warehouse Analysis')
        main_heading_format = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#4F81BD',
            'font_color': 'white'
        })

        sub_heading_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#9BC2E6',
            'font_color': 'black'
        })

        row_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter',
        })

        striped_row_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter',
            'bg_color': '#DDEBF7'
        })
        worksheet.merge_range('A1:E1', 'Warehouse Analysis', main_heading_format)
        headers = [
            "Company", "Partners", "Total Delay (Days)", "Total Cycle Time (Days)", "Total Product Quantity"
        ]
        worksheet.write_row(1, 0, headers, sub_heading_format)
        col_widths = [len(header) for header in headers]
        row_idx = 2
        for company_id, data in grouped_data.items():
            partners_list = "\n".join(set(data['partners']))

            values = [
                data['company_name'],
                partners_list,
                f"{data['total_delay']:.2f}",
                f"{data['total_cycle_time']:.2f}",
                str(data['total_product_qty'])
            ]
            row_format_to_use = row_format if row_idx % 2 == 0 else striped_row_format
            worksheet.write_row(row_idx, 0, values, row_format_to_use)
            for col_idx, value in enumerate(values):
                col_widths[col_idx] = max(col_widths[col_idx], len(value))
            row_idx += 1
        for col_idx, width in enumerate(col_widths):
            worksheet.set_column(col_idx, col_idx, width + 2)

