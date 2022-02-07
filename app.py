from flask import Flask
from flask import render_template, request, session, redirect, url_for
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
        session['current_problem_num'] = random.randint(1, 15)
        return redirect(url_for('index'))

    return render_template('index.html', problem = problem)
