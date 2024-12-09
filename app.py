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
        data = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, engine='openpyxl')
        return data.to_dict(orient="records")  # Возвращаем данные в виде списка словарей
    except Exception as e:
        print(f"Ошибка при чтении файла Excel: {e}")
        return []


def save_user_to_excel(username, email, password):
    user_data = {'username': username, 'email': email, 'password': password}
    df = pd.DataFrame([user_data])
    file_exists = os.path.isfile('data/users.xlsx')
    if file_exists:
        existing_df = pd.read_excel('data/users.xlsx', engine='openpyxl')
        df = pd.concat([existing_df, df], ignore_index=True)
    with pd.ExcelWriter('data/users.xlsx', engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False, header=True)


def user_exists(username):
    try:
        users = pd.read_excel('data/users.xlsx', engine='openpyxl').to_dict(orient='records')
        return any(user['username'] == username for user in users)
    except Exception as e:
        print(f"Ошибка при чтении файла Excel: {e}")
        return False


@app.route('/')
def index():
    if 'username' in session:
        return redirect(url_for('order_form'))
    return render_template('index.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        fullname = request.form['fullname']
        email = request.form['email']
        password = request.form['password']
        if user_exists(fullname):
            return 'Пользователь уже существует. Попробуйте другое имя.'
        hashed_password = generate_password_hash(password, method='pbkdf2:sha256')
        save_user_to_excel(fullname, email, hashed_password)
        return redirect(url_for('login'))
    return redirect(url_for('index'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        try:
            users = pd.read_excel('data/users.xlsx', engine='openpyxl').to_dict(orient='records')
            user = next((user for user in users if user['username'] == username), None)
            if user and check_password_hash(user['password'], password):
                session['username'] = user['username']
                return redirect(url_for('order_form'))
        except Exception as e:
            print(f"Ошибка при чтении файла Excel: {e}")
        return 'Неправильное имя пользователя или пароль.'
    return redirect(url_for('index'))


@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('index'))


@app.route('/order_form', methods=['GET', 'POST'])
def order_form():
    if 'username' not in session:
        return redirect(url_for('index'))
    variant1_data = read_excel("Вариант1")
    variant2_data = read_excel("Вариант2")
    return render_template("form.html", variant1_data=variant1_data, variant2_data=variant2_data)


@app.route('/load_products', methods=['POST'])
def load_products():
    variant = request.form.get('variant')
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
    file_exists = os.path.isfile('user_data.xlsx')
    if file_exists:
        existing_df = pd.read_excel('user_data.xlsx', engine='openpyxl')
        df = pd.concat([existing_df, df], ignore_index=True)
    with pd.ExcelWriter('user_data.xlsx', engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False, header=True)
    return jsonify({'message': 'Дані збережено!'})


if __name__ == "__main__":
    app.run(debug=True)
