import os
from dotenv import load_dotenv # type: ignore
import stripe # type: ignore
import gspread # type: ignore
from oauth2client.service_account import ServiceAccountCredentials # type: ignore
from datetime import datetime, timedelta
import math

def openSheet():
    CREDENTIALS_FILE = "./credentials.json";
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"];
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope);
    client = gspread.authorize(creds);

    SHEET_ID = "1lpdP6mc9RBsyPgEivzSsJCN2YmhQob-dqBGYOFczMxo";
    sheet = client.open_by_key(SHEET_ID).sheet1;
    return sheet;

def sendToSheets(sheet, row, cus_id, col_index):
    cus_ids = sheet.col_values(3);
    if col_index > sheet.col_count:
        sheet.add_cols(col_index - sheet.col_count);
        sheet.update_cell(1, col_index, str(datetime.now().year) + "-" + str(datetime.now().month));

    if cus_id in cus_ids:
        row_index = cus_ids.index(cus_id) + 1;
        new_mrr = row[-1];
        sheet.update_cell(row_index, col_index, new_mrr);
    else:
        sheet.append_row(row);

def calculateMRR(subscriptions, currency):
    active_mrr = {};
    for sub in subscriptions.auto_paging_iter():
        start_date = datetime.fromtimestamp(sub['start_date']);
        end_date = datetime.fromtimestamp(sub['ended_at']) if sub['ended_at'] else datetime.now();
        discount = sub['discount'];

        monthly_total = 0;
        for item in sub['items']['data']:
            price = item['price'];
            unit_amount = price['unit_amount'] / 100 if currency == "usd" else price['unit_amount'];
            interval_count = price['recurring']['interval_count'];
            quantity = item['quantity'];

            if price['recurring']['interval'] == 'month':
                monthly_amount = unit_amount * quantity * interval_count;
            elif price['recurring']['interval'] == 'year':
                monthly_amount = (unit_amount / 12) * quantity * interval_count;
            
            monthly_total += monthly_amount;

        if discount:
            percent_off = discount['coupon']['percent_off'];
            if percent_off:
                monthly_total = monthly_total * (1 - percent_off / 100);
        
            amount_off = discount['coupon']['amount_off'];
            if amount_off:
                monthly_total = monthly_total - (amount_off / 12);
        
        monthly_total = math.floor(monthly_total);
        current_month = start_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0);
        last_month = end_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0);
        while current_month <= last_month:
            key = current_month.strftime('%Y-%m');
            active_mrr[key] = format(round(active_mrr.get(key, 0) + monthly_total, 2), ".2f") if currency == 'usd' else active_mrr.get(key, 0) + monthly_total;
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

def getNewColumnIndex():
    start_year = 2021;
    start_month = 3;

    now = datetime.now();
    end_year = now.year;
    end_month = now.month;

    return ((end_year - start_year) * 12) + (end_month - start_month) + 7;

def main():
    load_dotenv();
    stripe.api_key = os.getenv('STRIPE_API_KEY');

    #Getting Customer Info Here
    sheet = openSheet();
    new_col_index = getNewColumnIndex();
    customers = stripe.Customer.list();

    for customer in customers.auto_paging_iter():        
        name = customer.name;
        email = customer.email;
        cus_id = customer.id;
        currency = customer.currency;

        subscriptions = stripe.Subscription.list(customer = cus_id, status='active', limit=100); 
        if not subscriptions.data:
            continue;

        start_date = datetime.fromtimestamp(subscriptions.data[0].start_date).strftime('%Y-%m-%d');
        end_date = "N/A" if subscriptions.data[0].ended_at == None else datetime.fromtimestamp(subscriptions.data[0].ended_at).strftime('%Y-%m-%d');
        mrr = calculateMRR(subscriptions, currency);

        row = [name, email, cus_id, start_date, end_date, currency] + mrr;
        sendToSheets(sheet, row, cus_id, new_col_index);

if __name__ == "__main__":
  main()