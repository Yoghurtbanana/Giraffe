from email.mime import image
from flask import Flask, render_template, request, flash, session
import openpyxl

app = Flask(__name__)
app.secret_key = 'ug@#gbG/n**FGDS'

wb = openpyxl.load_workbook('problems.xlsx')

def get_problem(problem_num):
    problemset = wb.active
    problem_dict = {
        'C': 'question', 'D': 'image_q', 'E': 'option_a', 'F': 'image_a',
        'G': 'option_b', 'H': 'image_b', 'I': 'option_c', 'J': 'image_c',
        'K': 'option_d', 'L': 'image_d'
    }
    problem = {'problem_num': problem_num}
    for i in problem_dict:
        problem[problem_dict[i]] = problemset[f'{i}{problem_num + 1}'].value
    return problem
    
@app.route('/', methods=['POST', 'GET'])
def index():
    current_problem_num = session.get('current_problem')
    if current_problem_num is None:
        current_problem_num = 1
        session['current_problem'] = current_problem_num
    problem = get_problem(current_problem_num)

    if request.method == 'POST':
        if 'options' not in request.form:
            flash('未選擇答案!')
        else:
            option = request.form['options']
            session['current_problem'] = current_problem_num + 1
            print(current_problem_num)
    return render_template('index.html', question = question, image_question = image_question,
    option_a = option_a, image_a = image_a, option_b = option_b, image_b = image_b,
    option_c = option_c, image_c = image_c, option_d = option_d, image_d = image_d)
    
    return render_template('index.html')

@app.route('/new', methods=['POST', 'GET'])
def new():
    if request.method == 'POST':
        question = request.form['question']
        option_a = request.form['option_a']
        option_b = request.form['option_b']
        option_c = request.form['option_c']
        option_d = request.form['option_d']
        if question and option_a and option_b and option_c and option_d:
            print('success')
        else:
            flash('有空格未填入!')
    return render_template('new.html')
