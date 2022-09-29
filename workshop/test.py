# from sqlalchemy import Column, Integer, String
# from flask import Flask
# from sqlalchemy import create_engine
# from flask_sqlalchemy import SQLAlchemy
from app import getCSV#
# engine = create_engine('sqlite:///sales.db', echo=True)
# from sqlalchemy.ext.declarative import declarative_base
# test = Flask(__name__)
# db = SQLAlchemy(test)
#
#
# class Customers(db.Model):
#     __tablename__ = 'customers'
#
#     id = Column(Integer, primary_key=True)
#     name = Column(String(100))
#     address = Column(String(100))
#     email = Column(String(100))
#
#
# from sqlalchemy.orm import sessionmaker
#
# Session = sessionmaker(bind=engine)
# session = Session()
#
# c1 = Customers(name='anurbek muhammed', address='Station Road Nanded', email='ravi@gmail.com')
# c2= Customers(name = 'azamat tarasbaev', address = 'Koti, Hyderabad', email = 'komal@gmail.com')
# session.add_all([c1, c2])
# session.commit()

Piar.query.filter(Piar.article_id.in_([13])).all()