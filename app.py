import random
import openpyxl
from flask import Flask
from flask import render_template, session, redirect, url_for

app = Flask(__name__)
app.secret_key = 'ug@#gbG/n**FGDS'

wb = openpyxl.load_workbook('problems.xlsx')
# 如有新增或刪除題目，請更新以下數字
MAX_PROBLEM_NUM = 15

def get_problem(problem_num):
    '''Gets the problem from a specified number'''
    problemset = wb.active
    if problemset[f'A{problem_num + 1}'].value is None:
        return None

    problem_dict = {
        'C': 'question', 'D': 'image_q', 'E': 'option_a', 'F': 'image_a',
        'G': 'option_b', 'H': 'image_b', 'I': 'option_c', 'J': 'image_c',
        'K': 'option_d', 'L': 'image_d', 'M': 'answer'
    }
    problem = {'problem_num': problem_num}
    for key, value in problem_dict.items():
        problem[value] = problemset[f'{key}{problem_num + 1}'].value
    return problem

@app.route('/', methods=['GET'])
def index():
    '''Main quiz page'''
    current_problem_num = session.get('current_problem_num')
    if current_problem_num is None:
        current_problem_num = random.randint(1, MAX_PROBLEM_NUM)
        session['current_problem_num'] = current_problem_num
    problem = get_problem(current_problem_num)

    return render_template('index.html', problem = problem)

@app.route('/next', methods=['GET'])
def next_problem():
    '''Retrieves a new random problem'''
    session['current_problem_num'] = random.randint(1, MAX_PROBLEM_NUM)
    return redirect(url_for('index'))
