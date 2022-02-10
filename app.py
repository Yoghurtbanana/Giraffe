import random
import openpyxl
from flask import Flask
from flask import render_template, session, redirect, url_for, request

app = Flask(__name__)
app.secret_key = 'ug@#gbG/n**FGDS'

wb = openpyxl.load_workbook('problems.xlsx')
# 如有新增或刪除題目，請更新以下數字
MAX_PROBLEM_NUM = 38

def get_problem(problem_num):
    '''Gets the problem from a specified number'''
    problemset = wb.active
    if problemset[f'A{problem_num + 1}'].value is None:
        return None

    problem_dict = {
        'C': 'question', 'D': 'image_q', 'E': 'option_a', 'F': 'image_a',
        'G': 'option_b', 'H': 'image_b', 'I': 'option_c', 'J': 'image_c',
        'K': 'option_d', 'L': 'image_d', 'M': 'answer', 'N': 'solution',
        'O': 'image_s'
    }
    problem = {'problem_num': problem_num}
    for key, value in problem_dict.items():
        problem[value] = problemset[f'{key}{problem_num + 1}'].value
    return problem

@app.route('/', methods=['GET'])
def index():
    '''Main page'''
    if request.args.get('problemId') is not None:
        problem_id = request.args.get('problemId')
        if problem_id.isnumeric():
            problem_id = int(problem_id)
            if problem_id > MAX_PROBLEM_NUM:
                problem_id = MAX_PROBLEM_NUM
            elif problem_id < 1:
                problem_id = 1
        else:
            problem_id = 1
        session['problem_num_main'] = problem_id

    problem_num_main = session.get('problem_num_main')
    if problem_num_main is None:
        problem_num_main = 1
        session['problem_num_main'] = problem_num_main
    problem = get_problem(problem_num_main)

    return render_template('index.html', problem = problem)

@app.route('/next', methods=['GET'])
def next_problem():
    '''Retrieves the next problem in order'''
    session['problem_num_main'] += 1
    if session['problem_num_main'] > MAX_PROBLEM_NUM:
        session['problem_num_main'] = 1
    return redirect(url_for('index'))

@app.route('/random', methods=['GET'])
def random_problem():
    '''Retrieves a random new problem'''
    session['problem_num_main'] = random.randint(1, MAX_PROBLEM_NUM)
    return redirect(url_for('index'))

@app.route('/practice', methods=['GET'])
def practice():
    '''Practice page'''
    problem_num_practice = session.get('problem_num_practice')
    if problem_num_practice is None:
        problem_num_practice = 1
        session['problem_num_practice'] = problem_num_practice
    problem = get_problem(problem_num_practice)
    return render_template('practice.html', problem = problem)

@app.route('/practice/next', methods=['GET'])
def new_problem():
    '''For practice page. Retrieves a random new problem'''
    session['problem_num_practice'] = random.randint(1, MAX_PROBLEM_NUM)
    return redirect(url_for('practice'))