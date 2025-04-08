import os
from dotenv import load_dotenv # type: ignore
import stripe # type: ignore
import gspread # type: ignore
from oauth2client.service_account import ServiceAccountCredentials # type: ignore
from datetime import datetime

def sendToSheets(rows):
    CREDENTIALS_FILE = "./credentials.json"
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)

    SHEET_ID = "1lpdP6mc9RBsyPgEivzSsJCN2YmhQob-dqBGYOFczMxo"
    sheet = client.open_by_key(SHEET_ID).sheet1

    sheet.append_rows(rows)
    print("Data successfully written!")

def main():
    load_dotenv();
    stripe.api_key = os.getenv('STRIPE_API_KEY');

    #Getting Customer Info Here
    rows = [];
    customers = stripe.Customer.list();

    for customer in customers.auto_paging_iter():        
        name = customer.name;
        email = customer.email;
        cus_id = customer.id;
        currency = customer.currency;

        subscription = stripe.Subscription.list(status = "all", customer = cus_id).data[0];
        start_date = datetime.fromtimestamp(subscription.start_date).strftime('%Y-%m-%d');
        end_date = "N/A" if subscription.canceled_at == None else datetime.fromtimestamp(subscription.canceled_at).strftime('%Y-%m-%d');

        customer = [name, email, cus_id, start_date, end_date, currency];
        rows.append(customer);

    sendToSheets(rows);

if __name__ == "__main__":
  main()