# https://docs.faculty.ai/user-guide/apis/flask_apis/flask_file_upload_download.html
import os
import datetime
import calendar
from dateutil.relativedelta import relativedelta
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from metrics.config import config

app = Flask(__name__)
app.config["DEBUG"] = True
cors = CORS(app)


@app.route('/', methods=['GET'])
def home():
    return "<h1>Dataverse Metrics</h1>"


@app.route('/api/metrics/list', methods=['GET'])
def get_reports_list():
    """Endpoint to list directories by date in archive base directory."""
    archive_dir = config.read('excel').get('archive_dir')
    print('REST archive_dir: ' + archive_dir)
    dirs = []
    for filename in os.listdir(archive_dir):
        path = os.path.join(archive_dir, filename)
        if os.path.isdir(path):
            dirs.append(filename)
    return jsonify(dirs)


@app.route('/api/metrics/report', methods=['GET'])
def get_report_by_date():
    """Download a file."""
    archive_dir = config.read('excel').get('archive_dir')
    metricsdate = request.args.get('date')
    if not metricsdate:
        metricsdate = get_latest_metrics_folder()
    filename = request.args.get('filename')
    if not filename:
        filename = 'Dataverse Usage Report.xlsx'
    file = os.path.join(archive_dir, metricsdate, filename)
    return send_from_directory(archive_dir, os.path.join(metricsdate, filename), as_attachment=True)


def get_latest_metrics_folder():
    _today = datetime.datetime.now()
    _start_date = (_today - relativedelta(years=1))
    # _start_date = (_today - relativedelta(years=1, months=1))
    start_date = datetime.date(_start_date.year, _start_date.month, 1).strftime('%Y-%m-%d')
    _end_date = (_today - relativedelta(months=1))
    # _end_date = (_today - relativedelta(months=2))
    end_date = datetime.date(_end_date.year, _end_date.month,
                             calendar.monthrange(_end_date.year, _end_date.month)[1]).strftime(
        '%Y-%m-%d')
    print('Generate Dataverse Usage Report for period: ' + start_date + ' to ' + end_date)
    month_sub_dir = os.path.join(start_date + ' to ' + end_date)
    print('Subdirectory: ' + month_sub_dir)
    return month_sub_dir


if __name__ == '__main__':
    app.run()
