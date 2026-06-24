from flask import Flask, render_template, request, send_file
from openpyxl import Workbook
from io import BytesIO


app = Flask(__name__)


# ==============================
# CALCULATION LOGIC
# ==============================

def calculate(age, wife_age, premium, payment_term, policy_term, rate):


    # AGE VALIDATION

    if age < 18 or age > 65:
        return {
            "ERROR":
            "Customer age must be between 18 and 65 years"
        }


    if wife_age < 18 or wife_age > 65:
        return {
            "ERROR":
            "Wife age must be between 18 and 65 years"
        }



    # ==============================
    # MATURITY CALCULATION
    # ==============================


    monthly_rate = rate / 12 / 100


    total_months = policy_term * 12


    premium_months = payment_term * 12


    maturity_amount = 0



    # Monthly premium compounding

    for month in range(premium_months):

        remaining_months = total_months - month


        maturity_amount += premium * (
            (1 + monthly_rate)
            ** remaining_months
        )




    # ==============================
    # 60% LUMPSUM
    # ==============================


    lump_sum_60 = maturity_amount * 0.60




    # ==============================
    # 40% PENSION CORPUS
    # ==============================


    pension_corpus = maturity_amount * 0.40




    # ==============================
    # PENSION @ 7%
    # ==============================


    annual_pension = pension_corpus * 0.07


    monthly_pension = annual_pension / 12




    # ==============================
    # CUSTOMER PENSION TILL AGE 85
    # ==============================


    maturity_age = age + policy_term

    customer_pension_years = 85 - maturity_age


    if customer_pension_years < 0:

        customer_pension_years = 0



    customer_total_pension = (
        monthly_pension *
        12 *
        customer_pension_years
    )




    # ==============================
    # WIFE PENSION AFTER POLICYHOLDER DEATH
    # ==============================


    # Difference between husband and wife age

    age_difference = age - wife_age



    # Policyholder assumed death at 85

    wife_age_at_husband_death = 85 - age_difference



    # Wife receives pension till her age 85

    wife_pension_years = 85 - wife_age_at_husband_death



    if wife_pension_years < 0:

        wife_pension_years = 0



    wife_total_pension = (

        monthly_pension *

        12 *

    wife_pension_years

)




    # ==============================
    # NOMINEE
    # ==============================


    nominee_amount = maturity_amount * 0.40




    total_benefit = (

        lump_sum_60

        +

        customer_total_pension

        +

        wife_total_pension

        +

        nominee_amount

    )





    return {


        "Policyholder Age":

        age,


        "Policyholder Wife Age":

        wife_age,



        "Maturity Amount (Minimum 12% ROR)":

        f"₹ {round(maturity_amount/10000000,2)} Crore",



        "60% Lump Sum Tax Free":

        f"₹ {round(lump_sum_60/10000000,2)} Crore",



        "40% Pension Corpus":

        f"₹ {round(pension_corpus/10000000,2)} Crore",



        "Annual Pension @ 7%":

        f"₹ {round(annual_pension)}",



        "Monthly Pension":

        f"₹ {round(monthly_pension)}",



        "Assumed Policyholder Life Expectency Till Age 85":

        f"₹ {round(customer_total_pension/10000000,2)} Crore",



        "Assumed Wife Life Expectency Till Age 85":

        f"₹ {round(wife_total_pension/100000,2)} Lakhs",



        "Nominee 40% Amount Tax Free lumpsum":

        f"₹ {round(nominee_amount/10000000,2)} Crore",



        "Total Benefit":

        f"₹ {round(total_benefit/10000000,2)} Crore"

    }




# ==============================
# WEBSITE
# ==============================


@app.route("/", methods=["GET","POST"])

def home():


    data = None



    if request.method == "POST":


        data = calculate(

            int(request.form["age"]),

            int(request.form["wife_age"]),

            int(request.form["premium"]),

            int(request.form["payment"]),

            int(request.form["policy"]),

            float(request.form["rate"])

        )



    return render_template(

        "index.html",

        data=data

    )





# ==============================
# EXCEL DOWNLOAD
# ==============================


@app.route("/excel", methods=["POST"])

def excel():


    result = calculate(

        int(request.form["age"]),

        int(request.form["wife_age"]),

        int(request.form["premium"]),

        int(request.form["payment"]),

        int(request.form["policy"]),

        float(request.form["rate"])

    )



    wb = Workbook()


    ws = wb.active


    ws.title = "Illustration"



    ws.append(

        [
            "Benefit",
            "Value"
        ]

    )



    for k,v in result.items():

        ws.append(

            [
                k,
                v
            ]

        )



    file = BytesIO()


    wb.save(file)


    file.seek(0)



    return send_file(

        file,

        as_attachment=True,

        download_name="Illustration.xlsx"

    )





if __name__ == "__main__":

    app.run(debug=True)