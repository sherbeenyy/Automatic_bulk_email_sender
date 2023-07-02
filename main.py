from email.message import EmailMessage
import pandas as pd
import ssl
import smtplib

path = r"E:\\automatic\\employees_data.xlsx"  # replace with your excel sheet path
all_data = pd.read_excel(path)


employees = {}

# converting each col to a list so we can manipulate data
names = list(all_data.Name)
phone_numbers = list(all_data.Phone_Number)
emails = list(all_data.Email)
salaries = list(all_data.Salary)

# putting them in a dict so we can acess them easily if further updates needed in the program
for i in range(len(names)):
    for name in names:
        for pn in phone_numbers:
            for email in emails:
                for salary in salaries:
                    employees[i] = {
                        "Name": name,
                        "phone_number": pn,
                        "Email": email,
                        "Salary": salary,
                    }


for i, name in enumerate(names):
    sender_email = "yourEmail"
    password = "yourPassword"
    recevier_email = f"{emails[i]}"

    subject = "Regarding your salary"
    body = f"hey, good morining {name} we hope you are doing well\n we just want to inform you that your salay is {salaries[i]}\n and you can come take it anytime from now...Thankyou for your hardwork"

    em = EmailMessage()
    em["From"] = sender_email
    em["To"] = recevier_email
    em["Subject"] = subject
    em.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as st:
        login = st.login(sender_email, password)
        # print(f"==>> login: {login}")

        send = st.sendmail(sender_email, recevier_email, em.as_string())
        # print(f"==>> send: {send}")
print("All emails sent")
