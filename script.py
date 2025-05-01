import stripe # type: ignore
import gspread # type: ignore
from oauth2client.service_account import ServiceAccountCredentials # type: ignore
from datetime import datetime, timedelta
import time
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

    col_index = ((end_year - start_year) * 12) + (end_month - start_month) + 7;
    if col_index > sheet.col_count:
        sheet.add_cols(col_index - sheet.col_count);
        sheet.update_cell(1, col_index, str(end_year) + "-" + str(end_month));

    return col_index;

def createNewCustomerRow():
    output = [];
    current = datetime(2021, 3, 1);
    now = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0);
    while current < now:
        output.append(0);
        current += timedelta(days=32);
        current = current.replace(day=1);
    
    return output;

def main():
    sheet = openSheet();
    new_col_index = addNewColumn(sheet);
    existing_customers = [c.strip().strip('"') for c in sheet.col_values(3)[1:]];
    new_row_zeros = createNewCustomerRow();

    today = datetime.today();
    today_str = today.strftime('%Y-%m-%d');
    first_day = today.replace(day=1);
    first_day_str = first_day.strftime('%Y-%m-%d');

    file_name = 'MRR_per_Subscriber_-_monthly_' + first_day_str + '_to_' + today_str +'.csv';
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

            if cus_id in existing_customers:
                row_index = existing_customers.index(cus_id) + 2;
                updates.append({
                    'range': gspread.utils.rowcol_to_a1(row_index, new_col_index),
                    'values': [[mrr]]
                });
            else:
                row = [name, email, cus_id, start_date, end_date, currency] + new_row_zeros + [mrr];
                new_rows.append(row);
        
        if updates:
            sheet.batch_update([{
                'range': update['range'],
                'values': update['values']
            } for update in updates])
        
        if new_rows:
            sheet.append_rows(new_rows);

if __name__ == "__main__":
  main()