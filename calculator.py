import json
from pathlib import Path
import pandas as pd
import streamlit as st
import io

import gspread
from oauth2client.service_account import ServiceAccountCredentials

GOOGLE_SHEET_RESPONSE_FORM = st.secrets["GOOGLE_SHEET_RESPONSE_FORM"]
HEADER_COLS = ["Email",	"Number of Returns", "Current Touch Points", "Ideal Touch Points", "Hourly Cost","Annual Cost", "Ideal Cost"]
SHEET_NAME = "Responses"


def calculate_annual_cost(num_returns, touch_points, hourly_cost):
    return (
        num_returns
        * touch_points
        * (hourly_cost * (st.session_state.time_to_email / 60))
    )

def main():

    if "time_to_email" not in st.session_state:
        st.session_state.time_to_email = 5

    with st.sidebar:
        st.subheader("How long does it take to send a personalized email in your firm?")
        st.write(
            "The default value is 5 minutes, but you can adjust this value to reflect the time it takes to send an email in your organization."
        )
        st.session_state.time_to_email = st.slider(
            "Time to send an email (minutes)", min_value=1, max_value=30, value=5
        )

    st.title("Tax Season Cost Calculator")

    if "num_returns" not in st.session_state:
        st.session_state.num_returns = 1000

    st.session_state.num_returns = st.number_input(
        "How many tax projects does your firm complete annually?",
        min_value=0,
        value=st.session_state.num_returns,
    )

    if "touch_points" not in st.session_state:
        st.session_state.touch_points = 3

    st.session_state.touch_points = st.slider(
        "How many personal updates are sent to your clients informing them of the status of their return?",
        min_value=0,
        max_value=50,
        value=st.session_state.touch_points,
    )

    if "ideal_touch_points" not in st.session_state:
        st.session_state.ideal_touch_points = 0

    st.session_state.ideal_touch_points = st.slider(
        "(Optional) If time and resources were not an obstacle, how many personal updates would you send them?",
        min_value=0,
        max_value=50,
        value=st.session_state.ideal_touch_points,
    )

    if "hourly_cost" not in st.session_state:
        st.session_state.hourly_cost = 23.0

    hourly_cost = st.number_input(
        "Enter the average hourly cost for each person sending the emails",
        min_value=0.0,
        value=st.session_state.hourly_cost,
    )

    annual_cost = calculate_annual_cost(
        st.session_state.num_returns,
        st.session_state.touch_points,
        st.session_state.hourly_cost,
    )
    # create a table that shows the cost of sending personalized emails, the cost of sending ideal personalized emails,
    # the number of hours annually that are spent sending personalized emails,
    # and the number of hours annually that are spent sending ideal personalized emails

    st.write("## Results")

    if 'ideal_cost' not in st.session_state:
        st.session_state.ideal_cost = 0

    if st.session_state.ideal_touch_points > 0:
        st.session_state.ideal_cost = calculate_annual_cost(
            st.session_state.num_returns,
            st.session_state.ideal_touch_points,
            hourly_cost,
        )

    monetary_df = pd.DataFrame(
        {
            "Metrics": [
                "Current Personalized Emails",
                "Ideal Personalized Emails",
            ],
            "Cost ($ USD)": [
                annual_cost,
                st.session_state.ideal_cost if st.session_state.ideal_touch_points > 0 else "N/A",
            ],
        }
    )

    time_df = pd.DataFrame(
        {
            "Metrics": [
                "Current Outreach Time (Hours)",
                "Ideal Outreach Time (Hours)",
            ],
            "Hours": [
                annual_cost / hourly_cost,
                st.session_state.ideal_cost / hourly_cost if st.session_state.ideal_touch_points > 0 else "N/A",
            ],
        }
    )

    # drop the index
    monetary_df.set_index("Metrics", inplace=True)
    time_df.set_index("Metrics", inplace=True)

    st.write("### Monetary Cost")
    st.table(monetary_df)

    st.write("### Time Cost")
    st.table(time_df)
    with st.expander("How did we calculate this?"):
        # write the formula
        st.write(
            f"Annual Cost = Number of Returns * Touch Points * (Hourly Cost * (Time to Send Email / 60))"
        )


    if max(st.session_state.ideal_cost, annual_cost) > 0:
        st.subheader(
            f"Would you like a free consultation to solve this ${max(st.session_state.ideal_cost, annual_cost):.2f} problem in your organization?"
        )
        email = st.text_input("Enter your email and we'll setup time with you!")
        if st.button("Submit"):
            st.write("Thank you! We will contact you soon.")

            # add the information to the google sheet
            # use the service account key to authenticate
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            # create a json file based on all of the information in the service account key
            # except that the data will come from the secrets file
            json_file_details = {
                "type": "service_account",
                "project_id": st.secrets["project_id"],
                "private_key_id": st.secrets["private_key_id"],
                "private_key": st.secrets["private_key"],
                "client_email": st.secrets["client_email"],
                "client_id": st.secrets["client_id"],
                "auth_uri": st.secrets["auth_uri"],
                "token_uri": st.secrets["token_uri"],
                "auth_provider_x509_cert_url": st.secrets["auth_provider_x509_cert_url"],
                "client_x509_cert_url": st.secrets["client_x509_cert_url"],
            }
            with open("pro-email-flow-48684d373739.json", "w") as json_file:
                json.dump(json_file_details, json_file)

            creds = ServiceAccountCredentials.from_json_keyfile_name(
                "pro-email-flow-48684d373739.json", scope
            )
            # remove the json file
            Path("pro-email-flow-48684d373739.json").unlink()
            client = gspread.authorize(creds)

            file_content: bytes = client.export(GOOGLE_SHEET_RESPONSE_FORM, format=gspread.client.ExportFormat.CSV)

            # given that file content is a <class 'bytes'> type object, we need to convert it to a string
            # and then use the StringIO class to convert it to a file-like object
            file_content = file_content.decode('utf-8')
            file_like_object = io.StringIO(file_content)
            data = pd.read_csv(file_like_object, names=HEADER_COLS, header=0)
            data.dropna(inplace=True)
            new_data = pd.DataFrame(
                {
                    "Email": email,
                    "Number of Returns": st.session_state.num_returns,
                    "Current Touch Points": st.session_state.touch_points,
                    "Ideal Touch Points": st.session_state.ideal_touch_points,
                    "Hourly Cost": hourly_cost,
                    "Annual Cost": annual_cost,
                    "Ideal Cost": st.session_state.ideal_cost,
                },
                index=[0],
            )
            data.loc[len(data)] = new_data.loc[0]
            data.to_csv(f"responses-{SHEET_NAME}.csv", index=False)
            client.import_csv(GOOGLE_SHEET_RESPONSE_FORM, Path(f"responses-{SHEET_NAME}.csv").open('r').read())
            Path(f"responses-{SHEET_NAME}.csv").unlink()


if __name__ == "__main__":
    main()
