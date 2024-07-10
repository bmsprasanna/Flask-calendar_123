from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from flask import Flask, redirect, url_for, render_template, session, request, flash, send_file
from flask_wtf import FlaskForm
from wtforms.fields import DateField, SubmitField
from wtforms.validators import DataRequired
from getpass import getuser
import win32api
import win32net
import pandas as pd
import json
import io
import pytz
# import xlsxwriter
def username():
    dc_name = win32net.NetGetAnyDCName()
    username = win32api.GetUserName()
    user_info = win32net.NetUserGetInfo(dc_name, username, 2)
    full_name = user_info["full_name"]
    return full_name

app = Flask(__name__)
app.config['SECRET_KEY'] = '#$%^&*'
app.config['SQLALCHEMY_DATABASE_URI'] = "sqlite:///todo.db"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

IST = pytz.timezone('Asia/Kolkata')


class InfoForm(FlaskForm):
    startdate = DateField('Select a Date', format='%Y-%m-%d', validators=[DataRequired()])
    submit = SubmitField('Submit')

class Todo(db.Model):
    sno = db.Column(db.Integer, primary_key=True)
    user_name = db.Column(db.String(200), nullable=False)
    group_lead = db.Column(db.String(500), nullable=False)
    project_name = db.Column(db.String(500), nullable=False)
    process_type = db.Column(db.String(500), nullable=False)
    count = db.Column(db.Integer, nullable=False)
    date_created = db.Column(db.DateTime, default=lambda: datetime.now(IST))
    date = db.Column(db.String(200), nullable=False)
    submitted_by = db.Column(db.String(200), nullable=False)
    comments = db.Column(db.String(500), nullable=False)

    def __repr__(self) -> str:
        return f"{self.sno} - {self.user_name}"

@app.route('/', methods=['GET', 'POST'])
def index():
    currentUser = username()
    form = InfoForm()
    if form.validate_on_submit():
        session['startdate'] = form.startdate.data.strftime('%Y-%m-%d')
        session['currentUser'] = currentUser
        return redirect(url_for('main_page'))

    return render_template('index.html', form=form, user_name='Date Selector', currentUser=currentUser)

@app.route('/main', methods=['GET', 'POST'])
def main_page():
    df = pd.read_excel('Process_Type.xlsx')
    df1 = pd.read_excel('Resource project wise list.xlsx')
    
    username_options = df1['Name'].dropna().tolist()
    # grouplead_options = df1['Group_Lead'].dropna().tolist()
    process_options = df['Process_Type'].dropna().tolist()

    startdate = session.get('startdate', datetime.utcnow().strftime('%Y-%m-%d'))
    grouplead_selected = {}
    date_selected = {}
    if startdate in df1.columns:
        for i, name in enumerate(df1['Name']):
            if pd.notna(name):
                date_selected[name] = str(df1[startdate].iloc[i])
    for i, name in enumerate(df1['Name']):
        if pd.notna(name):             
            grouplead_selected[name] = str(df1['Group_Lead'].iloc[i])

    print(date_selected)
    currentUser = username()

    if request.method == 'POST':
        try:
            user_name = request.form['user_name']
            group_lead = request.form['group_lead']
            process_type = request.form['process_type']
            count = request.form['count']
            comments = request.form['comments']
            project_name = request.form['project_name']

            if not all([user_name, group_lead, process_type, count, comments]):
                raise ValueError("One or more required fields are empty.")

            todo = Todo(user_name=user_name, group_lead=group_lead, project_name=project_name, 
                        process_type=process_type, count=count, 
                        date=startdate, submitted_by=currentUser, comments=comments)
            db.session.add(todo)
            db.session.commit()
            flash("Form submitted successfully!", "success")

        except Exception as e:
            flash(f"Error occurred: {str(e)}", "danger")

    allTodo = Todo.query.all()
    return render_template('main.html', allTodo=allTodo, username=currentUser, startdate=startdate, 
                           username_options=username_options, grouplead_selected=json.dumps(grouplead_selected), 
                           project_options=[], process_options=process_options, date_selected=json.dumps(date_selected))


@app.route('/update/<int:sno>', methods=['GET', 'POST'])
def update(sno):
    if request.method == 'POST':
        user_name = request.form['user_name']
        group_lead = request.form['group_lead']
        count = request.form['  ']
        todo = Todo.query.filter_by(sno=sno).first()
        todo.user_name = user_name
        todo.group_lead = group_lead
        todo.count = count
        db.session.add(todo)
        db.session.commit()
        return redirect("/main")

    todo = Todo.query.filter_by(sno=sno).first()
    return render_template('update.html', todo=todo)

@app.route('/delete/<int:sno>')
def delete(sno):
    todo = Todo.query.filter_by(sno=sno).first()
    db.session.delete(todo)
    db.session.commit()
    return redirect("/main")

# Command to create tables
@app.cli.command('init-db')
def init_db():
    db.create_all()
    print("Database initialized!")

@app.route('/export')
def export():
    allTodo = Todo.query.all()
    data = []
    for todo in allTodo:
        data.append({
            "SNo": todo.sno,
            "User Name": todo.user_name,
            "Group Lead": todo.group_lead,
            "Project Name": todo.project_name,
            "Process Type": todo.process_type,
            "Production Count": todo.count,
            "Date Created": todo.date_created,
            "Date": todo.date,
            "Submitted By": todo.submitted_by,
            "Comments": todo.comments
        })
    df = pd.DataFrame(data)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    output.seek(0)
    return send_file(output, download_name="production_report.xlsx", as_attachment=True)

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=8000)
