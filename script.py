import os
from dotenv import load_dotenv # type: ignore
import stripe # type: ignore
import gspread # type: ignore
from oauth2client.service_account import ServiceAccountCredentials # type: ignore
from datetime import datetime, timedelta
from decimal import Decimal

def sendToSheets(rows):
    CREDENTIALS_FILE = "./credentials.json"
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)

    SHEET_ID = "1lpdP6mc9RBsyPgEivzSsJCN2YmhQob-dqBGYOFczMxo"
    sheet = client.open_by_key(SHEET_ID).sheet1

    sheet.append_rows(rows)
    print("Data successfully written!")

def calculateMRR(subscriptions, currency):
    active_mrr = {};
    for sub in subscriptions.auto_paging_iter():
        start_date = datetime.fromtimestamp(sub['start_date'])
        end_date = datetime.fromtimestamp(sub['ended_at']) if sub['ended_at'] else datetime.now()

        for item in sub['items']['data']:
            price = item['price'];
            unit_amount = price['unit_amount'] / 100.00 if currency == "usd" else price['unit_amount'];
            interval_count = price['recurring']['interval_count'];
            quantity = item['quantity'];
            monthly_amount = round(unit_amount * quantity * interval_count, 2) if currency == 'usd' else unit_amount * quantity * interval_count;
            
            current_month = start_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0);
            last_month = end_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0);
            while current_month <= last_month:
                key = current_month.strftime('%Y-%m');
                active_mrr[key] = format(round(active_mrr.get(key, 0) + monthly_amount, 2), ".2f") if currency == 'usd' else active_mrr.get(key, 0) + monthly_amount;
                current_month += timedelta(days=32);
                current_month = current_month.replace(day=1);

    output = [];
    default_val = '0.00' if currency == 'usd' else 0;
    current = datetime(2021, 3, 1);
    now = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0);
    while current <= now:
        key = current.strftime('%Y-%m');
        output.append(active_mrr.get(key, default_val));
        current += timedelta(days=32);
        current = current.replace(day=1);

    return output;

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

        subscriptions = stripe.Subscription.list(customer = cus_id, status='all', limit=100);    
        start_date = datetime.fromtimestamp(subscriptions.data[0].start_date).strftime('%Y-%m-%d');
        end_date = "N/A" if subscriptions.data[0].ended_at == None else datetime.fromtimestamp(subscriptions.data[0].ended_at).strftime('%Y-%m-%d');
        mrr = calculateMRR(subscriptions, currency);

        row = [name, email, cus_id, start_date, end_date, currency] + mrr;
        rows.append(row);

    sendToSheets(rows);

if __name__ == "__main__":
  main()