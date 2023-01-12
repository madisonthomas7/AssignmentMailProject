import openpyxl

def getInfo():
    # Open the workbook
    wb = openpyxl.load_workbook('Assignments S23.xlsx')

    # Get the sheet
    sheet = wb['Sheet1']

    # Iterate through the rows
    for row in sheet.rows:
        # Process the row here
        due_date = row[2].value
        assignment = row[1].value
        courseNumber = row[0].value
        arrayOfAss = [courseNumber, assignment, due_date]
        if due_date == "Due date" or assignment == "Assignment Due" or courseNumber == "Course Number":
            continue;
        if due_date == None or assignment == None or courseNumber == None:
            break;
        print(arrayOfAss)

def sendEmail():
    from flask import Flask
    from flask_mail import Mail, Message
    app = Flask(__name__)
    mail = Mail(app)

    app.config['MAIL_SERVER'] = 'smtp.gmail.com'
    app.config['MAIL_PORT'] = 465
    app.config['MAIL_USERNAME'] = 'x@gmail.com'
    app.config['MAIL_PASSWORD'] = 'y'
    app.config['MAIL_USE_TLS'] = False
    app.config['MAIL_USE_SSL'] = True
    mail = Mail(app)

    @app.route("/")
    def index():
        msg = Message('Hello', sender='someone@gmail.com', recipients=['someone@gmail.com'])
        msg.body = "Hello Flask message sent from Flask-Mail"
        mail.send(msg)
        return "Sent"


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    getInfo();
    sendEmail()