from flask import Flask, redirect
from flask import render_template, request, flash, redirect, session, url_for
import openpyxl
import random

app = Flask(__name__)
app.secret_key = 'ug@#gbG/n**FGDS'

wb = openpyxl.load_workbook('problems.xlsx')

def get_problem(problem_num):
    problemset = wb.active
    if problemset[f'A{problem_num + 1}'].value == None:
        return None
    
    problem_dict = {
        'C': 'question', 'D': 'image_q', 'E': 'option_a', 'F': 'image_a',
        'G': 'option_b', 'H': 'image_b', 'I': 'option_c', 'J': 'image_c',
        'K': 'option_d', 'L': 'image_d', 'M': 'answer'
    }
    problem = {'problem_num': problem_num}
    for i in problem_dict:
        problem[problem_dict[i]] = problemset[f'{i}{problem_num + 1}'].value
    return problem
    
@app.route('/', methods=['POST', 'GET'])
def index():
    current_problem_num = session.get('current_problem_num')
    if current_problem_num == None:
        current_problem_num = random.randint(1, 15)
        session['current_problem_num'] = current_problem_num
    problem = get_problem(current_problem_num)

    if request.method == 'POST':
        if 'options' not in request.form:
            flash('未選擇答案!')
        else:
            option = request.form['options']
            if option == problem['answer']:
                flash('答案正確！')
            else:
                flash('答案錯誤！')
            session['current_problem_num'] = random.randint(1, 15)
            return redirect(url_for('index'))
            
    return render_template('index.html', question = problem['question'], image_q = problem['image_q'],
    option_a = problem['option_a'], image_a = problem['image_a'], option_b = problem['option_b'], image_b = problem['image_b'],
    option_c = problem['option_c'], image_c = problem['image_c'], option_d = problem['option_d'], image_d = problem['image_d'])

@app.route('/new', methods=['POST', 'GET'])
def new():
    if request.method == 'POST':
        question = request.form['question']
        option_a = request.form['option_a']
        option_b = request.form['option_b']
        option_c = request.form['option_c']
        option_d = request.form['option_d']
        if question and option_a and option_b and option_c and option_d:
            flash('成功！')
        else:
            flash('有空格未填入！')
    return render_template('new.html')
