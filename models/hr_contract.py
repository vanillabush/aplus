# -*- coding: utf-8 -*-

from odoo import fields,models

class PfaPensionReport(models.Model):
    _inherit = "hr.contract"

    employee_voluntary = fields.Float(string="Employee Voluntary")
    employer_voluntary = fields.Float(string="Employer Voluntary")