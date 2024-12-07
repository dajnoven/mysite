from flask import Flask, render_template, request, redirect, url_for, session, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
import os

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key'

EXCEL_FILE = "data/products.xlsx"  # Путь к файлу Excel


def read_excel(sheet_name):
    """Чтение данных из указанного листа Excel."""
    try:
        # Загружаем данные с указанного листа
        data = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        return data.to_dict(orient="records")  # Возвращаем данные в виде списка словарей
    except Exception as e:
        print(f"Ошибка при чтении файла Excel: {e}")
        return []


def save_user_to_excel(username, password):
    user_data = {'username': username, 'password': password}
    df = pd.DataFrame([user_data])
    # Проверяем, существует ли файл
    file_exists = os.path.isfile('users.xlsx')
    with pd.ExcelWriter('users.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, index=False, header=not file_exists)


@app.route('/')
def index():
    if 'username' in session:
        return redirect(url_for('order_form'))
    return render_template('index.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        hashed_password = generate_password_hash(password, method='sha256')
        # Сохранение данных пользователя в Excel файл
        save_user_to_excel(username, hashed_password)
        return redirect(url_for('login'))
    return render_template('register.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        # Чтение данных пользователей из Excel файла
        try:
            users = pd.read_excel('users.xlsx').to_dict(orient='records')
            user = next((user for user in users if user['username'] == username), None)
            if user and check_password_hash(user['password'], password):
                session['username'] = user['username']
                return redirect(url_for('order_form'))
        except Exception as e:
            print(f"Ошибка при чтении файла Excel: {e}")
        return 'Invalid username or password'
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('index'))


@app.route('/order_form', methods=['GET', 'POST'])
def order_form():
    if 'username' not in session:
        return redirect(url_for('login'))
    # При первой загрузке страницы загружаем данные для обоих вариантов
    variant1_data = read_excel("Вариант1")
    variant2_data = read_excel("Вариант2")
    return render_template("form.html", variant1_data=variant1_data, variant2_data=variant2_data)


@app.route('/load_products', methods=['POST'])
def load_products():
    variant = request.form.get('variant')  # Чтение параметра из формы
    # Загрузка данных в зависимости от варианта
    if variant == 'variant1':
        products = read_excel("Вариант1")
    elif variant == 'variant2':
        products = read_excel("Вариант2")
    else:
        products = []
    return jsonify({'products': products})


@app.route('/save_user_data', methods=['POST'])
def save_user_data():
    user_data = request.form.to_dict()
    df = pd.DataFrame([user_data])
    # Проверяем, существует ли файл
    file_exists = os.path.isfile('user_data.xlsx')
    with pd.ExcelWriter('user_data.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, index=False, header=not file_exists)
    return jsonify({'message': 'Дані збережено!'})


if __name__ == "__main__":
    app.run(debug=True)
