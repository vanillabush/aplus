# -*- coding: utf-8 -*-
{
    'name': "APLUS",
    'version': '18.0.1.0',
    'depends': ['base','stock','sale_management','hr','hr_contract','hr_payroll'],
    'author': "Author Name",
    'category': 'Category',
    'license': 'LGPL-3',
    'description': """ 
     This module contains following reports:
      Client Stock Report,
      Custom Inventory Report,
      Employee Paye Report,
      Inventory Held Report,
      Payment Schedule Report,
      PFA Pension Report,
      Payroll Schedule Report
     """,
  
    'data': [
        'security/ir.model.access.csv',
        
        'views/stock_picking.xml',
        'views/hr_employee.xml',
        'views/hr_contract.xml',
        
        'wizard/stock_management_report_wizard_view.xml',
      ],
    'install_requires':['openpyxl'],
    'auto_install': False,
    'installable': True,
    'application': True,
}