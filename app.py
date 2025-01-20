from flask import Flask, jsonify, render_template, request, redirect, url_for, flash, session
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Cambiar por un valor seguro en producción

# Ruta al archivo Excel
EXCEL_FILE = "users.xlsx"

# Crear archivo Excel si no existe
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=["email", "password"])
    df.to_excel(EXCEL_FILE, index=False)

def load_users():
    """Cargar los usuarios desde el archivo Excel."""
    try:
        return pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        # Si el archivo no existe, retornar un DataFrame vacío
        return pd.DataFrame(columns=["email", "password", "username", "remail_id"])


import pandas as pd

EXCEL_FILE = 'users.xlsx'  # Nombre del archivo Excel

def load_users():
    """Cargar los usuarios desde el archivo Excel."""
    try:
        return pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        # Si el archivo no existe, retornar un DataFrame vacío
        return pd.DataFrame(columns=["email", "password", "username", "remail_id"])

def save_user(email, password, username):
    """Guardar un nuevo usuario con un ID único en el archivo Excel."""
    # Cargar los usuarios existentes
    df = load_users()

    # Generar el nuevo ID en formato 're-X', donde X es el siguiente número
    if df.empty:
        remail_id = 're-1'  # Si no hay usuarios, comenzamos con 're-1'
    else:
        last_id = df['remail_id'].iloc[-1]  # Obtener el último ID
        last_num = int(last_id.split('-')[1])  # Extraer el número del último ID
        remail_id = f're-{last_num + 1}'  # Incrementar el número para el nuevo ID

    # Crear el nuevo usuario
    new_user = pd.DataFrame([{
        "email": email, 
        "password": password, 
        "username": username, 
        "remail_id": remail_id
    }])

    # Añadir el nuevo usuario al DataFrame
    df = pd.concat([df, new_user], ignore_index=True)

    # Guardar el DataFrame actualizado en el archivo Excel
    # Especifica el engine explícitamente
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')



@app.route('/')
def landing_page():
    return render_template('landing.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        users = load_users()

        # Asegurarse de que las columnas sean de tipo string
        users['email'] = users['email'].astype(str)
        users['password'] = users['password'].astype(str)

        # Verificar las credenciales
        user_match = users[(users['email'].str.strip() == email.strip()) & (users['password'].str.strip() == password.strip())]
        if not user_match.empty:
            username = user_match['username'].iloc[0]
            remail_id = user_match['remail_id'].iloc[0]
            status = user_match['status'].iloc[0]
            session['user'] = {"email": email, "username": username, "remail_id":remail_id, "status":status}
            flash('Login exitoso', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Correo o contraseña incorrectos', 'danger')

    return render_template('login.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form['email']
        username = request.form['username']
        password = request.form['password']
        users = load_users()
        if (users['email'] == email).any():
            flash('El correo ya está registrado.', 'warning')
        else:
            save_user(email, password, username)
            flash('Registro exitoso. Por favor, inicia sesión.', 'success')
            return redirect(url_for('login'))
    return render_template('register.html')

@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        flash('Por favor, inicia sesión primero.', 'warning')
        return redirect(url_for('login'))
    
    df = pd.read_excel("últimos_100_correos_rapido.xlsx")
    emails = df.to_dict(orient="records")
    print("sesion", session['user'])
    return render_template('dashboard.html', user=session['user'], emails=emails)

@app.route('/update_email_status', methods=['POST'])
def update_email_status():
    data = request.json
    email_id = data['email_id']
    new_status = data['new_status']

    # Cargar el archivo Excel
    df = pd.read_excel('últimos_100_correos_rapido.xlsx')

    # Actualizar el estado del correo
    df.loc[df['Msg_ID'] == email_id, 'Estado'] = new_status

    # Guardar los cambios
    df.to_excel('últimos_100_correos_rapido.xlsx', index=False)

    return jsonify({"message": "Estado actualizado correctamente"}), 200

@app.route('/add_status', methods=['POST'])
def add_status():
    try:
        # Obtener datos del request
        data = request.json
        new_status = data.get('status', '').strip()

        # Validar el estado
        if not new_status:
            return jsonify({'error': 'Invalid status'}), 400

        # Leer el archivo Excel
        users_df = pd.read_excel('users.xlsx')

        # Actualizar la columna 'status' para todos los usuarios
        for index, row in users_df.iterrows():
            current_statuses = eval(row['status']) if isinstance(row['status'], str) else []
            if new_status not in current_statuses:
                current_statuses.append(new_status)
            users_df.at[index, 'status'] = str(current_statuses)

        # Guardar el Excel actualizado
        users_df.to_excel('users.xlsx', index=False)

        return jsonify({'message': 'Status added successfully'}), 200

    except Exception as e:
        print('Error:', e)
        return jsonify({'error': 'An error occurred while updating the Excel file.'}), 500

@app.route('/logout')
def logout():
    session.pop('user', None)
    flash('Sesión cerrada exitosamente.', 'success')
    return redirect(url_for('landing_page'))

if __name__ == '__main__':
    app.run(debug=True)
