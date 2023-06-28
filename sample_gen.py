from prepare_invoice import create_invoice
from prepare_invoice import all_sheets
from prepare_invoice import Transactions
from prepare_invoice import Transaction

df = all_sheets['transactions']
all_transactions = Transactions([Transaction(row) for index, row in df.iterrows()])
all_transactions.filter_out_type('Payment')
all_transactions.filter_out_processed()
all_transactions.filter_by_funding_source('Acme')
create_invoice(all_transactions, invoice_name = "upwork_expenses_acme", invoice_from = "Coyote", 
               invoice_to = "ACME")

