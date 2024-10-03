from odoo import models, fields, api
from odoo.exceptions import ValidationError

class WarehouseAnalysisWizard(models.TransientModel):
    _name = 'warehouse.analysis.wizard'
    _description = "Warehouse Analysis Report Wizard"

    start_date = fields.Date(string='Start Date', default=fields.Date.today)  
    end_date = fields.Date(string='End Date')

    @api.constrains('start_date', 'end_date')
    def _check_dates(self):
        for record in self:
            if record.start_date and record.end_date and record.end_date < record.start_date:
                raise ValidationError('End Date must be greater than Start Date.')

    def action_warehouse_analysis_report(self):
        data = {
            'start': self.start_date,
            'end': self.end_date
        }
        return self.env.ref('warehouse_analysis_report.action_warehouse_analysis_excel_report').report_action(self, data=data)


