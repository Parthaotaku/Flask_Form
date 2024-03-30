from bson import ObjectId
from flask import Flask, render_template, request, session, flash,  make_response
from pymongo import MongoClient
from functools import wraps
import io
from openpyxl import Workbook

mongo_uri = "mongodb+srv://ashgrey1818:lgcpNAxHYdIQa38b@records.qk5uqvc.mongodb.net/?retryWrites=true&w=majority&appName=records"
client = MongoClient(mongo_uri)

db = client["Student"]
collection = db["Form"]

db2 = client['user']
users_collection = db2['admin']
app = Flask(__name__)
app.secret_key = "blood"


def check_credentials(username, password):
    user = users_collection.find_one({'username': username, 'password': password})
    return user is not None


@app.route('/login', methods=['GET', 'POST'])
def login():
    username = request.form['first']
    password = request.form['password']
    form_data = collection.find()
    # Check if the username and password match
    user = users_collection.find_one({'password': password, 'username': username})
    if user:
        return render_template('dashboard.html', form_data=form_data)

    else:
        return render_template('login.html', message='* Invalid credentials! Please try again *')


@app.route('/', methods=["GET", "POST"])
def index():  # put application's code here
    return render_template("login.html")


@app.route("/submit", methods=["GET", "POST"])
def submit():
    fname = request.form["fullname"]
    roll = request.form["roll"]
    reg = request.form["reg"]
    school = request.form["school"]
    prog = request.form["prog"]
    email = request.form["email"]
    phone = request.form["phone"]
    gender = request.form["gender"]
    city = request.form["city"]
    state = request.form["state"]
    pincode = request.form["pincode"]

    # Insert document into MongoDB collection
    collection.insert_one({
        "Full_Name": fname,
        "Roll_No": roll,
        "Registration_No": reg,
        "School": school,
        "Program": prog,
        "email": email,
        "phone": phone,
        "Gender": gender,
        "address": {
            "city": city,
            "state": state,
            "pincode": pincode
        }
    })
    return render_template("ok.html")


@app.route('/update', methods=['POST'])
def update_document():
    form_data = collection.find()
    if request.method == 'POST':
        document_id = request.form.get('selected_document_id')

        if document_id:
            # Retrieve form data without concatenating with the document ID
            full_name = request.form['Full_Name_' + document_id]
            roll_no = request.form['Roll_No_' + document_id]
            registration_no = request.form['Registration_No_' + document_id]
            school = request.form['School_' + document_id]
            program = request.form['Program_' + document_id]
            email = request.form['email_' + document_id]
            phone = request.form['phone_' + document_id]
            gender = request.form['Gender_' + document_id]
            city = request.form['city_' + document_id]
            state = request.form['state_' + document_id]
            pincode = request.form['pincode_' + document_id]

            # Update the document in MongoDB
            collection.update_one(
                {"_id": ObjectId(document_id)},
                {"$set": {
                    "Full_Name": full_name,
                    "Roll_No": roll_no,
                    "Registration_No": registration_no,
                    "School": school,
                    "Program": program,
                    "email": email,
                    "phone": phone,
                    "Gender": gender,
                    "address.city": city,
                    "address.state": state,
                    "address.pincode": pincode
                }}
            )

            return render_template('dashboard.html', form_data=form_data)

    return render_template('dashboard.html', form_data=form_data,
                           message='* Record not selected! Please Select the record before saving *')


@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'):
        return render_template('login.html')
    form_data = collection.find()
    return render_template('dashboard.html', form_data=form_data)


# Logout route
@app.route('/logout', methods=['GET', 'POST'])
def logout():
    session.pop('logged_in', None)
    return render_template('login.html')


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            return render_template('login.html')
        return f(*args, **kwargs)

    return decorated_function


# Example usage of login_required decorator
@app.route('/restricted')
@login_required
def restricted_area():
    return 'You are in a restricted area!'


@app.route("/form", methods=["GET", "POST"])
def form_page():
    return render_template("index.html")


@app.route('/download_excel', methods=['GET'])
def download_excel():
    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active
    form_data = collection.find()
    # Add some data to the worksheet

    ws.append(['Full Name', 'Roll No', 'Registration No', 'School', 'Program', 'Email', 'Phone', 'Gender', 'City', 'State', 'Pincode'])
    # Assuming you have form_data containing the form submissions
    for data in form_data:
        city = data.get('address', {}).get('city', '')
        state = data.get('address', {}).get('state', '')
        pincode = data.get('address', {}).get('pincode', '')
        ws.append([data.get('Full_Name'), data.get('Roll_No'), data.get('Registration_No'), data.get('School'), data.get('Program'), data.get('email'), data.get('phone'), data.get('Gender'), city, state, pincode])

    # Save the workbook to a BytesIO object
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    # Create a response with the Excel file as attachment
    response = make_response(output.getvalue())
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response.headers['Content-Disposition'] = 'attachment; filename=form_data.xlsx'
    return response


if __name__ == '__main__':
    app.run(debug=False,host='0.0.0.0')
