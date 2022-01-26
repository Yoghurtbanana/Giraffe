from email.mime import image
from flask import Flask, render_template, request, flash, session
import openpyxl

app = Flask(__name__)
app.secret_key = 'ug@#gbG/n**FGDS'

@app.route('/', methods=['POST', 'GET'])
def index(question = None, image_question = None, option_a = None, image_a = None, option_b = None, image_b = None, option_c = None, image_c = None, option_d = None, image_d = None):
    error = None
    wb = openpyxl.load_workbook('problems.xlsx')
    problemset = wb.active
    current_problem = session.get('current_problem')
    if current_problem is None:
        current_problem = 1
        session['current_problem'] = current_problem

    current_problem = 1
    question = problemset['C' + str(current_problem + 1)].value
    image_question = problemset['D' + str(current_problem + 1)].value
    option_a = problemset['E' + str(current_problem + 1)].value
    image_a = problemset['F' + str(current_problem + 1)].value
    option_b = problemset['G' + str(current_problem + 1)].value
    image_b = problemset['H' + str(current_problem + 1)].value
    option_c = problemset['I' + str(current_problem + 1)].value
    image_c = problemset['J' + str(current_problem + 1)].value
    option_d = problemset['K' + str(current_problem + 1)].value
    image_d = problemset['L' + str(current_problem + 1)].value

    if request.method == 'POST':
        if 'options' not in request.form:
            error = '未選擇答案!'
            flash(error)
        else:
            option = request.form['options']
            session['current_problem'] = current_problem + 1
            print(current_problem)
    
    return render_template('index.html', question = question, image_question = image_question,
    option_a = option_a, image_a = image_a, option_b = option_b, image_b = image_b,
    option_c = option_c, image_c = image_c, option_d = option_d, image_d = image_d,
    error = error)

@app.route('/new', methods=['POST', 'GET'])
def new():
    error = None
    if request.method == 'POST':
        question = request.form['question']
        option_a = request.form['option_a']
        option_b = request.form['option_b']
        option_c = request.form['option_c']
        option_d = request.form['option_d']
        if question and option_a and option_b and option_c and option_d:
            print('success')
        else:
            error = '有空格未填入!'
            flash(error)
    return render_template('new.html', error = error)
