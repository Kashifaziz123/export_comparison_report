# -*- coding: utf-8 -*-
from datetime import datetime
from odoo import models, fields, api


class StockReportXls(models.AbstractModel):
    _name = 'report.export_comparison_report.comparison_report_xls.xlsx'
    _inherit = 'report.report_xlsx.abstract'

    def get_lines(self):
        lines = {}
        lines['a'] = 1
        return lines

    def generate_xlsx_report(self, workbook, data, lines):
        period_arr = []
        count = 1
        for line_period in lines.period_preview:
            duration = line_period.name.split('-')
            from_date = duration[0] + '-' + duration[1] + '-' + duration[2]
            print('.............', from_date)
            to_date = duration[3] + '-' + duration[4] + '-' + duration[5]
            print('.............', to_date)
            print('............', )
            period_dict = {}
            period_dict['period'] = 'Period ' + str(count)
            print('........period', line_period.name)
            records_forecast = self.env['forecast.product'].search(
                [('forecast_id', '=', lines.sale_forecast.id), ('period_start_date', '=', from_date),
                 ('period_end_date', '=', to_date)])
            productarr = []
            print('records_forecast...............', records_forecast)
            for records in records_forecast:
                product = {}
                product['product'] = records.product_id.name
                product['forecast'] = records.forecast_qty
                probable_qty = 0
                crm_lead = self.env['crm.lead'].search(
                    [('date_deadline', '>=', from_date), ('date_deadline', '<=', to_date)])
                print('..............crmlead', crm_lead)
                for crm_lead_line in crm_lead:
                    print('crm........', crm_lead_line.probability)
                    if crm_lead_line.probability:
                        for line_expected_demand_product in crm_lead_line.expected_demand_product_ids:
                            if line_expected_demand_product.product.id == records.product_id.id:
                                probable_qty += line_expected_demand_product.quantity * (
                                        crm_lead_line.probability / 100)
                product['probable_quantity'] = round(probable_qty, 2)
                productarr.append(product)
            period_dict['products'] = productarr
            period_arr.append(period_dict)

            count += 1
        print('..................arrays', period_arr)
        sheet = workbook.add_worksheet("Comparison Report")
        sheet.set_zoom(75)
        cell_heading_format = workbook.add_format(
            {'bold': True, 'align': 'center', "font_name": "Arial", "font_size": 16})
        cell_43_format = workbook.add_format({'bold': True, "font_name": "Arial", "font_size": 12, 'border': 1})
        cell_period_format = workbook.add_format(
            {'bold': True, "font_name": "Arial", "font_size": 12, 'align': 'center', 'border': 1})
        sheet.merge_range('E1:F3', 'Forecasted vs Projected Sale', cell_heading_format)

        sheet.write(4, 3,
                    'Product',
                    cell_43_format)
        row_period = 3
        col_period = 4
        sheet.set_column("D:AD", 30)
        is_product_done = False
        if not lines.combine_forecast:
            for period_arr_line in period_arr:
                sheet.merge_range(row_period, col_period, row_period, col_period + 1,
                                  period_arr_line['period'],
                                  cell_period_format)
                sheet.write(row_period + 1, col_period,
                            'Forecasted',
                            cell_period_format)
                sheet.write(row_period + 1, col_period + 1,
                            'Projected',
                            cell_period_format)
                product_row = row_period + 2
                for product in period_arr_line['products']:
                    if not is_product_done:
                        sheet.write(product_row, col_period - 1,
                                    product['product'],
                                    cell_43_format)
                    sheet.write(product_row, col_period,
                                product['forecast'],
                                cell_43_format)
                    sheet.write(product_row, col_period + 1,
                                product['probable_quantity'],
                                cell_43_format)
                    product_row += 1
                is_product_done = True
                col_period += 2
        else:
            sheet.merge_range(row_period, col_period, row_period, col_period + 1,
                              'Results',
                              cell_period_format)
            sheet.write(row_period + 1, col_period,
                        'Forecasted',
                        cell_period_format)
            sheet.write(row_period + 1, col_period + 1,
                        'Projected',
                        cell_period_format)
            forecast_list = []
            probable_qty_list = []

            for period_arr_line in period_arr:
                product_row = row_period + 2
                count = 0
                for product in period_arr_line['products']:
                    if not is_product_done:
                        sheet.write(product_row, col_period - 1,
                                    product['product'],
                                    cell_43_format)
                        print('..........', float(product['forecast']))
                        print(product['forecast'])
                        probable_qty_list.append(float(product['probable_quantity']))
                        forecast_list.append(float(product['forecast']))
                    else:
                        forecast_list[count] = forecast_list[count] + float(product['forecast'])
                        probable_qty_list[count] = probable_qty_list[count] + float(product['probable_quantity'])
                    count += 1
                    product_row += 1
                is_product_done = True
                print('forecast;ist.........', forecast_list)
            row_new = row_period + 2
            for val in forecast_list:
                sheet.write(row_new, col_period,
                            val,
                            cell_period_format)
                row_new += 1
            row_new_2=row_period+2
            for val in probable_qty_list:
                sheet.write(row_new_2, col_period+1,
                            val,
                            cell_period_format)
                row_new_2 += 1


class StockReport(models.TransientModel):
    _name = "comparison.report"
    _description = "Forecasted Vs Projected Sale"

    sale_forecast = fields.Many2one('sale.forecast', string='Select SaleForeCast ', required=True)
    period = fields.Char('Period', compute='onchange_depends_saleforecast', store=True)
    period_count = fields.Integer('No. of Periods', compute='onchange_depends_saleforecast', store=True)
    start_date = fields.Date(string='Start Date', compute='onchange_depends_saleforecast', store=True)
    period_preview = fields.Many2many('sale.forecast.periods')
    combine_forecast = fields.Boolean('Combine Forecast')

    @api.depends('sale_forecast')
    def onchange_depends_saleforecast(self):
        if self.sale_forecast:
            sale_forecast_records = self.env['sale.forecast'].search([('name', '=', self.sale_forecast.name)])
            if sale_forecast_records:
                self.period = sale_forecast_records.period
                self.period_count = sale_forecast_records.period_count
                self.start_date = sale_forecast_records.start_date

    @api.onchange('sale_forecast')
    def onchange_saleforecast(self):
        if self.sale_forecast:
            sale_forecast_records = self.env['sale.forecast'].search([('name', '=', self.sale_forecast.name)])
            if sale_forecast_records:
                periods = self.env['sale.forecast.periods'].search([('period_id', '=', self.sale_forecast.id)])
                if periods:
                    mlist = []
                    dupplicate_arr = []
                    for val in periods:
                        if val.name not in dupplicate_arr:
                            mlist.append(val.id)
                        dupplicate_arr.append(val.name)
                    print('..............', dupplicate_arr)
                    print('................', mlist)
                    return {'domain': {'period_preview': [('id', 'in', mlist)]}}
                else:
                    return {'domain': {'period_preview': [('id', '=', 0)]}}
        else:
            return {'domain': {'period_preview': [('id', '=', 0)]}}

    @api.multi
    def export_report(self):
        datas = {}
        return self.env.ref('export_comparison_report.comparison_xlsx').report_action(self)
