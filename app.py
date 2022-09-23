from flask import Flask, render_template, url_for, request, redirect, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import openpyxl as exl
from flask_migrate import Migrate
from sqlalchemy.orm import relationship, backref
from sqlalchemy import Column, Integer, String, Text, DateTime, ForeignKey
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

engine = create_engine('sqlite:///database.db', echo=True)



wb = exl.load_workbook(filename="bsg.xlsx")
ws = wb.active
sheet = wb["BSG-PR"]


app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
# migrate = Migrate(app, db)



class Article(db.Model):
    __tablename__ = 'article'
    id = Column(Integer, primary_key=True)
    requis = Column(String(50))
    vendor = Column(Text, nullable=True)
    desc = Column(Text)
    invoice = Column(Text)
    date = Column(DateTime, default=datetime.utcnow)
    piar = relationship('Piar', backref='article', lazy='dynamic')




class Piar(db.Model):
    __tablename__ = 'piar'
    id = Column(Integer, primary_key=True)
    unit_price = Column(Integer)
    unit_desc = Column(Text)
    purpose = Column(Text, nullable=True)
    quality = Column(Integer)
    respon = Column(String(100), nullable=True)
    cost = Column(Text, nullable=True)
    article_id = Column(Integer, ForeignKey('article.id'))



@app.route('/index')
@app.route('/')
def index():
    articles = Article.query.order_by(Article.id.desc()).all()
    return render_template('index.html', articles=articles)

Session = sessionmaker(bind=engine)
session = Session()

@app.route('/create', methods=['POST', 'GET'])
def create():
    if request.method == 'POST':
        # form for piar index
        id = request.form.get('id')
        # total_price = request.form['total_price']
        desc = request.form['desc']
        vendor = request.form['vendor']
        date = request.form.get('date')
        requis = request.form['requis']
        invoice = request.form['invoice']
        currency = request.form['currency']
        # total_price2 = f'{total_price} {currency}'

        try:
            # form for purchsing
            unit_desc = request.form['unit_desc']
            unit_price = request.form['unit_price']
            # unit_total_price = request.form['unit_total_price']
            purpose = request.form['purpose']
            quality = request.form['quality']
            cost = request.form['cost']
            respon = request.form['respon']


            unit_desc1 = request.form['unit_desc1']
            unit_price1 = request.form['unit_price1']
            # unit_total_price = request.form['unit_total_price']
            purpose1 = request.form['purpose1']
            quality1 = request.form['quality1']
            cost1 = request.form['cost1']
            respon1 = request.form['respon1']

        except:
            return 'error add piar'


        article = Article(id=id,
                          invoice=invoice,
                          requis=requis,
                          # total_price=total_price2,
                          desc=desc,
                          vendor=vendor,
                          date=date)
        try:
            piar = Piar(unit_price=unit_price,
                        unit_desc=unit_desc,
                        # unit_total_price=unit_total_price,
                        purpose=purpose,
                        respon=respon,
                        quality=quality,
                        cost=cost,
                        article=article
                        )
            piar1 = Piar(unit_price=unit_price1,
                        unit_desc=unit_desc1,
                        # unit_total_price=unit_total_price,
                        purpose=purpose1,
                        respon=respon1,
                        quality=quality1,
                        cost=cost1,
                        article=article
                        )
        except:
            return 'error with piara'


        try:
            session.add(article)
            session.add_all([piar, piar1])
            session.commit()
            return redirect('/index')
        except:
            return 'error database'
    else:
        articles = Article.query.all()
        piars = Piar.query.all()



        cost_center = ['Workshop equipment | Оборудование для цеха',
                       'Workshop tools | Инструменты для цеха',
                       'Workshop consumables | Расходные материалы для цеха',
                       'PPE, Safety equipment |СИЗ и оборудование по ТБ',
                       'Office equipment and furniture | Офисное оборудование и мебель',
                       'Office consumables and stationary | Расходные материалы и канцелярские товары ',
                       'Purchase of services for workshop | Услуги для цеха',
                       'Car maintenance and car consumables | Обслуживание автомобилей и расходные материалы к автомобилям',
                       'Office maintenance (household expenses, tea, cofffee) | Обеспечение офиса',
                       'Repair of fixed assets | Ремонт основных средств',
                       'Third-party services | Услуги третьей стороны',
                       'Trainings | Тренинги',
                       'IT software applications, licenses, etc',
                       ]



        names = ['Aizada Abtay',
                 'Anuarbek Muhammed',
                 'Akhmediyar Salykov',
                 'Adrian Owen',
                 'Aizada Muratova',
                 'Andrew Davidson',
                 'Berik Gibbatulin',
                 'Dias Kuraishov',
                 'Lyailya Sarbayeva',
                 'Nurlan Zhunussov',
                 'John Carter',
                 'Akmaral Izteleuova',
                 'Adilet Kumarov',
                 'Nurgul Mukanova',
                 'Salamat Maukenov',
                 'Tarasbaev Azamat',
                 'Sergey Mazur',
                 'Karlygash Lepessova',
                 'Nurgul Zhubanova',
                 'Aimira Dzhumagalieva',
                 'Shynar Ramazanova']

        return render_template('create.html',  names=names, articles=articles, piars=piars)

@app.route('/index/<int:id>', methods=['GET', 'POST'])
def getCSV(id):
    abc = [id]
    value = Article.query.filter(Article.id.in_(abc)).all()
    for xl in value:
        ws['E4'] = 'SVR-PR-'+'{:05}'.format(xl.id)
        ws['F19'] = xl.name
        ws['I4'] = xl.supp
        ws['C19'] = xl.text
        ws['N19'] = xl.price
        ws['E5'] = xl.date
        wb.save('test.xlsx')

    return send_file(download_name='save.xlsx', path_or_file='test.xlsx')


@app.route('/test')
def test():
    return render_template('test.html')

@app.route('/search')
def search_piar():
    params = {name: request.args.get(name) for name in ['id', 'text', 'name']}
    params = {k: v for k, v in params.items() if v} # отфильтровываем пустые параметры
    book = Article.query.filter_by(**params).first_or_404()
    return render_template('show_book.html', book=book)

if __name__=='__main__':
    app.run(debug=True, port='8080')
