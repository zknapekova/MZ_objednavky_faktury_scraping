import pandas as pd
from flask import Flask, render_template, request, send_file
from mysql_config import objednavky_db_connection_cloud as db_connection
from sqlalchemy import create_engine


app = Flask(__name__)
engine = create_engine(
    f"mysql+pymysql://{db_connection['user']}:{db_connection['password']}@{db_connection['host']}/{db_connection['database']}?charset=utf8mb4",
    connect_args={'ssl': {'ssl_ca': '/etc/ssl/cert.pem'}})

def get_data(keywords, price_min=None, price_max=None):
    with engine.connect() as conn:

        price_min_criteria = ""
        price_max_criteria = ""
        if price_min:
            price_min_criteria = f" and cena > {price_min} "
        if price_max:
            price_max_criteria = f" and cena < {price_max} "
        query = "select objednavatel, cena, datum, popis, link from priame_objednavky where popis like '%%"+keywords+"%%'"+ price_min_criteria + price_max_criteria +" order by cena desc, datum desc limit 1000"
        result = pd.read_sql(query, con=conn)
    return result

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        price_min = request.form.get('price_filter_min')
        price_max = request.form.get('price_filter_max')
        data = get_data(keywords=request.form.get('search'), price_min=price_min, price_max=price_max)
        if 'export_to_csv' in request.form:
            data.to_csv('exported_data.csv', index=False)
            return send_file('exported_data.csv', attachment_filename='exported_data.csv', as_attachment=True)
    else:
        data = pd.DataFrame()
    return render_template('index.html', table=data.to_html())


if __name__ == '__main__':
    app.run(debug=True)

