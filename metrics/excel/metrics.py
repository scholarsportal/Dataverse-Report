import os
import sys
import shutil
import calendar
import datetime
from dateutil.relativedelta import relativedelta
from distutils.util import strtobool
from openpyxl import load_workbook

from metrics.config import config
from metrics.email import metricsemail
from metrics.excel import datasetdownloads as dtsheet
from metrics.excel import dataversedownloads as dvsheet
from metrics.excel import filetypes as ftsheet
from metrics.excel import subjects as sbsheet
from metrics.excel import useraffiliation as uasheet


class Metrics:

    @staticmethod
    def generate():
        history = config.read('history')
        if strtobool(history.get('regenerate')):
            Metrics.generate_historical_reports(history)
        else:
            Metrics.generate_latest_report()

    @staticmethod
    def generate_historical_reports(history):
        _today = datetime.datetime.now()
        history_start_date = datetime.datetime.strptime(history.get('startdate'), '%Y-%m-%d')
        periods = int(history.get('number_of_time_periods'))
        if not history_start_date or not periods:
            print("Please update metrics configuration (metrics.cfg) 'history' section.")
            sys.exit(1)
        while history_start_date + relativedelta(months=1) < _today:
            _start_date = (history_start_date - relativedelta(years=1))
            _end_date = (history_start_date - relativedelta(months=1))
            start_date = datetime.date(_start_date.year, _start_date.month, 1).strftime('%Y-%m-%d')
            end_date = datetime.date(_end_date.year, _end_date.month,
                                     calendar.monthrange(_end_date.year, _end_date.month)[1]).strftime(
                '%Y-%m-%d')
            print('Generating Dataverse Historical Usage Report for period: ' + start_date + ' to ' + end_date)
            history_start_date = history_start_date + relativedelta(months=1)
            periods = periods - 1
            Metrics.create(start_date, end_date)

    @staticmethod
    def generate_latest_report():
        _today = datetime.datetime.now()
        _start_date = (_today - relativedelta(years=1))
        _end_date = (_today - relativedelta(months=1))
        start_date = datetime.date(_start_date.year, _start_date.month, 1).strftime('%Y-%m-%d')
        end_date = datetime.date(_end_date.year, _end_date.month,
                                 calendar.monthrange(_end_date.year, _end_date.month)[1]).strftime(
            '%Y-%m-%d')
        print('Generate Dataverse Usage Report for period: ' + start_date + ' to ' + end_date)
        Metrics.create(start_date, end_date)

    @staticmethod
    def create(start_date, end_date):
        month_sub_dir = Metrics.create_month_sub_dir(start_date, end_date)
        workbook = Metrics.create_workbook(month_sub_dir)

        print('Processing sheet 1')
        dvsheet.DataverseDownloads(start_date, end_date, workbook, month_sub_dir).create()
        print('Processing sheet 2')
        dtsheet.DatasetDownloads(start_date, end_date, workbook, month_sub_dir).create()
        print('Processing sheet 3')
        ftsheet.FileTypes(start_date, end_date, workbook, month_sub_dir).create()
        print('Processing sheet 4')
        sbsheet.Subjects(start_date, end_date, workbook, month_sub_dir).create()
        print('Processing sheet 5')
        uasheet.UserAffiliation(start_date, end_date, workbook, month_sub_dir).create()

        metricsemail.send_email()

    @staticmethod
    def create_workbook(month_sub_dir):
        usage_report = config.read('excel').get('report')
        template = config.read('excel').get('template')
        config_dir = config.read('configuration').get('config_dir')
        shutil.copyfile(os.path.join(config_dir, template), os.path.join(month_sub_dir, usage_report))
        return load_workbook(os.path.join(month_sub_dir, usage_report), data_only=True)

    @staticmethod
    def create_month_sub_dir(start_date, end_date):
        archive_dir = config.read('excel').get('archive_dir')
        if not os.path.isdir(archive_dir):
            os.mkdir(archive_dir)
        month_sub_dir = os.path.join(archive_dir, start_date + ' to ' + end_date)
        print('Creating subdirectory: ' + month_sub_dir)
        if os.path.exists(month_sub_dir):
            shutil.rmtree(month_sub_dir)
        os.mkdir(month_sub_dir)
        return month_sub_dir
