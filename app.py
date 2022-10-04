import io
from flask import Flask, render_template, url_for, request, redirect, send_file, send_from_directory,Response
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
# from flask_migrate import Migrate
from sqlalchemy.orm import relationship, backref
from sqlalchemy import Column, Integer, String, Text, DateTime, ForeignKey
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import openpyxl as exl
from openpyxl.styles import Font ,Alignment,Border,Side
import os
from tempfile import NamedTemporaryFile

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
# migrate = Migrate(app, db)
engine = create_engine('sqlite:///database.db', echo=True)


Session = sessionmaker(bind=engine)
session = Session()


class Article(db.Model):
    __tablename__ = 'article'
    id = Column(Integer, primary_key=True)
    requis = Column(String(50))
    vendor = Column(Text, nullable=False)
    desc = Column(Text, nullable=False)
    invoice = Column(Text)
    date = Column(DateTime, default=datetime.utcnow)
    piar = relationship('Piar', backref='article', lazy='dynamic')


class Piar(db.Model):
    __tablename__ = 'piar'
    id = Column(Integer, primary_key=True)
    unit_price = Column(Integer, nullable=False)
    unit_desc = Column(Text, nullable=False)
    purpose = Column(Text, nullable=False)
    quality = Column(Integer, nullable=False)
    respon = Column(String(100), nullable=False)
    cost = Column(Text, nullable=False)
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
        # currency = request.form['currency']
        # total_price2 = f'{total_price} {currency}'


        article = Article(id=id,
                          invoice=invoice,
                          requis=requis,
                          # total_price=total_price2,
                          desc=desc,
                          vendor=vendor,
                          date=date)




        session.add(article)
        session.commit()
        return redirect('/index')

    else:
        articles = Article.query.all()



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

        return render_template('create.html',  names=names, articles=articles)


@app.route('/index/<int:id>/pr', methods=['POST', 'GET'])
def getPR(id):
    if request.method == "POST":
        unit_price = request.form['unit_price']
        unit_desc = request.form['unit_desc']
        purpose = request.form['purpose']
        quality = request.form['quality']
        respon = request.form['respon']
        cost = request.form['cost']

        piar = Piar(
            unit_price=unit_price,
            unit_desc=unit_desc,
            purpose=purpose,
            quality=quality,
            respon=respon,
            cost=cost,
            article_id=id
        )


        session.add(piar)
        session.commit()
        return redirect(f'/index/{id}/pr')

    else:
        piars = Piar.query.filter(Piar.article_id.in_([id])).all()

        return render_template('pr.html', piars=piars)

# config openpyxl
wb = exl.load_workbook(filename="test.xlsx")
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

        for xl in value:
            ws['E4'] = 'SVR-PR-' + '{:05}'.format(xl.id)
            ws['I19'] = xl.vendor
            ws['I4'] = xl.requis
            ws['E5'] = xl.date
            save_name = 'SVR-PR-' + '{:05}'.format(xl.id)

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


        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, download_name=f"{save_name}.xlsx", as_attachment=True, max_age=0)

if __name__=='__main__':
    app.run(debug=True)