{
    'name': 'Warehouse Analysis Excel Reports',
    'version': '17.0.1.0',
    'summary': 'Warehouse analysis excel report',
    'author': 'Shins',
    'license': 'LGPL-3',
    'depends': ['stock','report_xlsx'],
    'data': [
        'security/ir.model.access.csv',
        # 'views/warehouse_analysis_menu.xml',
        'wizard/warehouse_analysis_wizard_view.xml',
        'report/warehouse_analysis_excel_report_views.xml',
    ],
    'installable': True,
    'application': False,
}
