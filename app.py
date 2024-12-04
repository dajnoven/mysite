from flask import Flask, render_template, request, jsonify
import pandas as pd

app = Flask(__name__)

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


@app.route("/", methods=["GET", "POST"])
def order_form():
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


if __name__ == "__main__":
    app.run(debug=True)
