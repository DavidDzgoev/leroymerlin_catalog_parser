from flask import send_file, Flask, request, after_this_request
from parser import get_excel_from_category
import os

app = Flask(__name__)


@app.route('/get_excel/<cat>')
def get_xlsx(cat):
    get_excel_from_category(cat)

    @after_this_request
    def remove_file(response):
        try:
            os.remove(f'{cat}.xlsx')
        except Exception as error:
            app.logger.error("Error removing or closing downloaded file handle", error)
        return response

    return send_file(f'{cat}.xlsx', mimetype='xlsx')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
