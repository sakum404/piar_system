from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField, TextAreaField
from wtforms.validators import DataRequired, Email

class LoginForm(FlaskForm):
    surname = StringField('Username', validators=[DataRequired()])
    lastname = StringField('Lastname', validators=[DataRequired()])
    submit_flask = SubmitField('Submit')
