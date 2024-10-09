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
        query = """
            SELECT
                parent.id as partner_id,
                parent.name as company_name,
                SUM(sp.delay) as total_delay,
                SUM(sp.cycle_time) as total_cycle_time,
                SUM(sp.product_qty) as total_product_qty,
                parent.company_id
            FROM
                res_partner rp
            LEFT JOIN res_partner parent ON rp.parent_id = parent.id OR rp.id = parent.id
            LEFT JOIN (
                SELECT
                    sp.partner_id,
                    (EXTRACT(EPOCH FROM AVG(sp.date_done - sp.scheduled_date)) / (24 * 60 * 60))::decimal(16,2) as delay,
                    (EXTRACT(EPOCH FROM AVG(sp.date_done - sp.date)) / (24 * 60 * 60))::decimal(16,2) as cycle_time,
                    SUM(sm.product_qty) as product_qty
                FROM
                    stock_move sm
                LEFT JOIN stock_picking sp ON sm.picking_id = sp.id
                WHERE
                    sm.date >= %s AND sm.date <= %s
                GROUP BY
                    sp.partner_id
            ) sp ON sp.partner_id = rp.id
            WHERE
                parent.is_company = True
            GROUP BY
                parent.id, parent.name, parent.company_id
            ORDER BY
                parent.name
        """
        self.env.cr.execute(query, (start_date, end_date))
        results = self.env.cr.fetchall()
        grouped_data = defaultdict(lambda: {
            'company_name': '',
            'total_delay': 0,
            'total_cycle_time': 0,
            'total_product_qty': 0,
        })
    
        for row in results:
            partner_id, company_name, total_delay, total_cycle_time, total_product_qty, company_id = row
            grouped_data[partner_id]['company_name'] = company_name or 'No Company'
            grouped_data[partner_id]['total_delay'] += total_delay or 0
            grouped_data[partner_id]['total_cycle_time'] += total_cycle_time or 0
            grouped_data[partner_id]['total_product_qty'] += total_product_qty or 0
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
        row_format = workbook.add_format({'font_size': 10, 'align': 'left', 'valign': 'vcenter'})
        striped_row_format = workbook.add_format({'font_size': 10, 'align': 'left', 'valign': 'vcenter', 'bg_color': '#DDEBF7'})
        worksheet.merge_range('A1:D1', 'Warehouse Analysis', main_heading_format)
        headers = ["Company", "Total Delay (Days)", "Total Cycle Time (Days)", "Total Product Quantity"]
        worksheet.write_row(1, 0, headers, sub_heading_format)

        col_widths = [len(header) for header in headers]
        row_idx = 2
        for company_id, data in grouped_data.items():
            values = [
                data['company_name'],
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
