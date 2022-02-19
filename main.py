from flask import send_file, Flask, render_template, after_this_request
from parser import get_excel_from_category
import os

app = Flask(__name__, template_folder="templates")


@app.route('/leroymerlin_parser')
def root():
    return render_template('index.html')


@app.route('/leroymerlin_parser/get_excel/<cat>')
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
    port = int(os.environ.get('PORT', 80))
    app.run(host='0.0.0.0', port=port)
