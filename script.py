import os
from dotenv import load_dotenv # type: ignore
import stripe # type: ignore
import gspread # type: ignore
from oauth2client.service_account import ServiceAccountCredentials # type: ignore
from datetime import datetime, timedelta
import csv

def openSheet():
    CREDENTIALS_FILE = "./credentials.json";
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"];
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope);
    client = gspread.authorize(creds);

    SHEET_ID = "1lpdP6mc9RBsyPgEivzSsJCN2YmhQob-dqBGYOFczMxo";
    sheet = client.open_by_key(SHEET_ID).sheet1;
    return sheet;

def addNewColumn(sheet):
    start_year = 2021;
    start_month = 3;

    now = datetime.now();
    end_year = now.year;
    end_month = now.month;

    if now.day < 7:
        if now.month == 1:
            end_month = 12;
        else:
            end_month = now.month - 1;

    col_index = ((end_year - start_year) * 12) + (end_month - start_month) + 9;
    if col_index > sheet.col_count:
        sheet.add_cols(col_index - sheet.col_count);
        sheet.update_cell(1, col_index, str(end_year) + "-" + str(end_month));

    return col_index;

def createNewCustomerRow():
    output = [];
    current = datetime(2021, 3, 1);
    now = datetime.now();

    if now.day < 7:
        if now.month == 1:
            now = now.replace(year = now.year - 1, month = 12, day = 1, hour = 0, minute = 0, second = 0, microsecond = 0);
        else:
            now = now.replace(month = now.month - 1, day = 1, hour = 0, minute = 0, second = 0, microsecond = 0);
    else:
        now = now.replace(day=1, hour = 0, minute = 0, second = 0, microsecond = 0);

    while current < now:
        output.append(0);
        current += timedelta(days=32);
        current = current.replace(day=1);
    
    return output;

def getReportName():
    today = datetime.today();
    if today.day < 7:
        if today.month == 1:
            start_date = today.replace(year = today.year - 1, month = 12, day = 1);
        else:
            start_date = today.replace(month = today.month - 1, day = 1);
    else:
        start_date = today.replace(day = 1);

    first_day_str = start_date.strftime('%Y-%m-%d');
    today_str = today.strftime('%Y-%m-%d');
    return 'mrr_reports/MRR_per_Subscriber_-_monthly_' + first_day_str + '_to_' + today_str +'.csv';

def getSubscriptions():
    subscriptions = {};
    has_more = True;
    starting_after = None;

    while has_more:
        response = stripe.Subscription.list(status='active', limit = 100, starting_after = starting_after);
        for sub in response['data']:
            customer_id = sub['customer'];
            subscriptions[customer_id] = {
                'current_period_start': sub['current_period_start'],
                'interval': sub['items']['data'][0]['price']['recurring']['interval']
            }

        has_more = response['has_more'];
        if response['data']:
            starting_after = response['data'][-1]['id'];

    return subscriptions;

def main():
    load_dotenv();
    stripe.api_key = os.getenv('STRIPE_API_KEY');

    sheet = openSheet();
    new_col_index = addNewColumn(sheet);
    existing_customers = [c.strip().strip('"') for c in sheet.col_values(3)[1:]];
    existing_end_dates = [e.strip().strip('"') for e in sheet.col_values(5)[1:]];
    customer_enddate_map = {key : value for key, value in zip(existing_customers, existing_end_dates)};
    new_row_zeros = createNewCustomerRow();
    file_name = getReportName();
    subscriptions = getSubscriptions();

    with open(file_name, 'r', encoding='utf-8') as file:
        customers = csv.reader(file)
        next(customers)

        updates = [];
        new_rows = [];
        for customer in customers:
            name = customer[0];
            email = customer[1];
            cus_id = customer[2];
            start_date = customer[3];
            end_date = customer[4];
            currency = customer[5];
            mrr = customer[6];
            
            subscription = subscriptions.get(cus_id);
            if subscription:
                recent_payment = datetime.fromtimestamp(subscription['current_period_start']).strftime('%Y-%m-%d');
                interval = subscription['interval'];
            else:
                recent_payment = 'N/A';
                interval = 'N/A';

            if cus_id in existing_customers:
                row_index = existing_customers.index(cus_id) + 2;
                updates.append({'range': gspread.utils.rowcol_to_a1(row_index, new_col_index),'values': [[mrr]]});
                updates.append({'range': gspread.utils.rowcol_to_a1(row_index, 6),'values': [[recent_payment]]});
                updates.append({'range': gspread.utils.rowcol_to_a1(row_index, 7),'values': [[interval]]});
                if end_date != customer_enddate_map.get(cus_id):
                    updates.append({'range': gspread.utils.rowcol_to_a1(row_index, 5),'values': [[end_date]]});
            
            else:
                row = [name, email, cus_id, start_date, end_date, recent_payment, interval, currency] + new_row_zeros + [mrr];
                new_rows.append(row);
        
        if updates:
            sheet.batch_update([{
                'range': update['range'],
                'values': update['values']
            } for update in updates])
        
        if new_rows:
            sheet.append_rows(new_rows);

    print("Writing MRR to Google Sheet has been completed");
if __name__ == "__main__":
  main()