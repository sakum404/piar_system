from flask import Flask, render_template, url_for, request, redirect, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from flask_migrate import Migrate
from sqlalchemy.orm import relationship, backref
from sqlalchemy import Column, Integer, String, Text, DateTime, ForeignKey
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import openpyxl as exl
from openpyxl.styles import Font ,Alignment,Border,Side

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
migrate = Migrate(app, db)
engine = create_engine('sqlite:///database.db', echo=True)

Session = sessionmaker(bind=engine)
session = Session()

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

# config openpyxl
wb = exl.load_workbook(filename="bsg.xlsx")
ws = wb.active
sheet = wb["BSG-PR"]
font_style = Font(name='Verdana', sz='11')
alig_style_left = Alignment(wrapText=True, horizontal='left', vertical='center')
alig_style_right = Alignment(wrapText=True, horizontal='right', vertical='center')
alig_style_center = Alignment(wrapText=True, horizontal='center', vertical='center')
top = Side(border_style='thin')
bottom = Side(border_style='thin')
left = Side(border_style='thin')
right = Side(border_style='thin')
border = Border(top=top, left=left, right=right, bottom=bottom)
border_bottom_top = Border(bottom=Side(border_style='thin'), top=Side(border_style='thin'))
nextrow = sheet.max_row + 1
@app.route('/index/<int:id>', methods=['GET', 'POST'])
def getCSV(id):
        abc = [id]
        value = Article.query.filter(Article.id.in_(abc)).all()
        piar_value = Piar.query.filter(Piar.article_id.in_(abc)).all()

        for i, lx in enumerate(piar_value):
            ws.merge_cells(f'C{19+i}:D{19+i}')
            ws.merge_cells(f'F{19 + i}:G{19 + i}')
            ws.merge_cells(f'J{19 + i}:K{19 + i}')
            ws[f'D{19 + i}'].border = border_bottom_top
            ws[f'G{19 + i}'].border = border_bottom_top
            ws[f'K{19 + i}'].border = border_bottom_top

            ws[f'B{19 + i}'].border = border
            ws[f'B{19 + i}'].value = 1+i
            ws[f'B{19 + i}'].font = font_style
            ws.row_dimensions[20 + i].height = 71
            ws[f'B{19 + i}'].alignment = alig_style_center

            ws[f'C{19+i}'].border = border
            ws[f'C{19+i}'].value = lx.unit_desc
            ws[f'C{19 + i}'].font = font_style
            ws.row_dimensions[20+i].height = 71
            ws[f'C{19 + i}'].alignment = alig_style_center

            ws[f'E{19 + i}'].border = border
            ws[f'E{19 + i}'].value = lx.quality
            ws[f'E{19 + i}'].font = font_style
            ws.row_dimensions[20 + i].height = 71
            ws[f'E{19 + i}'].alignment = alig_style_center

            ws[f'F{19 + i}'].border = border
            ws[f'F{19 + i}'].value = lx.respon
            ws[f'F{19 + i}'].font = font_style
            ws.row_dimensions[20 + i].height = 71
            ws[f'F{19 + i}'].alignment = alig_style_center

            ws[f'L{19 + i}'].border = border
            ws[f'L{19 + i}'].value = lx.cost
            ws[f'L{19 + i}'].font = font_style
            ws.row_dimensions[20 + i].height = 71
            ws[f'L{19 + i}'].alignment = alig_style_center

            ws[f'M{19 + i}'].border = border
            ws[f'M{19 + i}'].value = lx.unit_price
            ws[f'M{19 + i}'].font = font_style
            ws.row_dimensions[20 + i].height = 71
            ws[f'M{19 + i}'].alignment = alig_style_center

        for xl in value:
            ws['E4'] = 'SVR-PR-' + '{:05}'.format(xl.id)
            ws['I19'] = xl.vendor
            ws[f'I{nextrow}'] = xl.vendor
            ws['I4'] = xl.requis
            ws['E5'] = xl.date

        wb.save('test.xlsx')

        return send_file(download_name='save.xlsx', path_or_file='test.xlsx')

if __name__=='__main__':
    app.run(debug=True, port='8080')