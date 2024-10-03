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

        print(f"Fetched {len(reports)} reports between {start_date} and {end_date}")
        grouped_data = defaultdict(lambda: {'partner_name': '', 'delay': 0, 'cycle_time': 0, 'product_qty': 0, 'count': 0})
        for report in reports:
            partner_id = report.partner_id.id if report.partner_id else 'No Partner'
            partner_name = report.partner_id.name if report.partner_id else 'No Partner'
            grouped_data[partner_id]['partner_name'] = partner_name
            grouped_data[partner_id]['delay'] += report.delay or 0
            grouped_data[partner_id]['cycle_time'] += report.cycle_time or 0
            grouped_data[partner_id]['product_qty'] += report.product_qty or 0
            grouped_data[partner_id]['count'] += 1
        all_partners = self.env['res.partner'].search([])
        for partner in all_partners:
            if partner.id not in grouped_data:
                grouped_data[partner.id] = {
                    'partner_name': partner.name,
                    'delay': 0,
                    'cycle_time': 0,
                    'product_qty': 0,
                    'count': 0
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
        worksheet.merge_range('A1:D1', 'Warehouse Analysis', main_heading_format)
        headers = [
            "Partner", "Total Delay (Days)", "Total Cycle Time (Days)", "Total Product Quantity"
        ]
        worksheet.write_row(1, 0, headers, sub_heading_format)
        col_widths = [len(header) for header in headers]
        row_idx = 2
        for partner_id, data in grouped_data.items():
            average_delay = data['delay'] / data['count'] if data['count'] else 0
            average_cycle_time = data['cycle_time'] / data['count'] if data['count'] else 0
            values = [
                data['partner_name'],
                f"{average_delay:.2f}",
                f"{average_cycle_time:.2f}",
                str(data['product_qty'])
            ]
            row_format_to_use = row_format if row_idx % 2 == 0 else striped_row_format
            worksheet.write_row(row_idx, 0, values, row_format_to_use)
            for col_idx, value in enumerate(values):
                col_widths[col_idx] = max(col_widths[col_idx], len(value))
            row_idx += 1
        for col_idx, width in enumerate(col_widths):
            worksheet.set_column(col_idx, col_idx, width + 2)

