from flask import Flask, request, redirect, render_template, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import pandas as pd # need to install to test
from io import BytesIO # built-in in python, no need to install
import xlsxwriter # need to install to test
from werkzeug import secure_filename

app = Flask(__name__)
app.config['DEBUG'] = True

app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://import-export:123@localhost:8889/import-export'
app.config['SQLALCHEMY_ECHO'] = True 


db = SQLAlchemy(app) 
app.secret_key = '@!@##Adfasfcvdg12'


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    first_name = db.Column(db.String(120))
    last_name = db.Column(db.String(120))
    pin = db.Column(db.Integer)
    cohort = db.Column(db.Integer)
    city = db.Column(db.String(120))
    attendance_date = db.relationship("Attendance", backref="owner")

    def __init__(self, first_name, last_name, pin=0000, cohort=1, city="Miami"):
        self.first_name = first_name
        self.last_name = last_name
        self.pin = pin
        self.cohort = cohort
        self.city = city

class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date_now = db.Column(db.Date)
    time_now = db.Column(db.Time)
    present = db.Column(db.Boolean)
    owner_id = db.Column(db.Integer, db.ForeignKey("student.id"))
    

    def __init__(self, owner, date_now=None, time_now=None):
        if date_now == None:
            date_now = date.today()
        if time_now == None:
            time_now = datetime.time(datetime.now()).strftime("%H:%M:%S")
        self.time_now = time_now
        self.date_now = date_now
        self.owner = owner
        self.present = False


@app.route('/', methods = ['GET'])
def index():
    return render_template('index.html', title = 'Test')


@app.route('/upload_file', methods = ['POST'])
def upload_file():

    if request.method == 'POST':
        # TODO add validation, in case user didn't select any file to upload
        file = request.files['file']

    # ----------- Reads Files and pushes to student table -------------
    df = pd.read_excel(file)
    # df.columns is a list of all the table headings, 'First Name' and 'Last Name'
    # in this case.
    first_name = list(df[df.columns[0]])
    last_name = list(df[df.columns[1]])
    #  names is a list of tupples in the form of (first_name, last_name)
    names = zip(first_name,last_name)

    # creates a record for row in students.xlsx into the student table.
    for name in names:
        student = Student(name[0].title(),name[1].title())
        db.session.add(student)
    db.session.commit()
    return render_template('index.html', title = 'Test')




@app.route('/download_list')
def download_list():
        
        students = Student.query.all()
        first_names = []
        last_names = []
        
        # Get the information from student table and populates the arrays above
        for student in students:
            first_names.append(student.first_name)
            last_names.append(student.last_name)

        # creates a dictionary where names will be the headers for the spreadsheet
        # and the values(lists, see above) are rows for each column.
        df = pd.DataFrame({'First Name': first_names, 'Last Name': last_names})
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine = 'xlsxwriter')
        df.to_excel(writer, 'Sheet1', index=False)
        writer.save()
        output.seek(0)

        return send_file(output, attachment_filename='attendance_list.xlsx',as_attachment=True)

if __name__ == '__main__':
    app.run()
