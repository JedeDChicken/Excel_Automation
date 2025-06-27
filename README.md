# Automated Discount application and Chart construction, and Test Excel generation, and Concurrency

*Automating Excel spreadsheets to process thousands of spreadsheets in under a second is very useful especially in today's era of big data. This is a simple implementation of automation using Python, particularly of adding discounts and charts for a price column in multiple Excel files and sheets, in parallel. This program also includes a function for generating a test Excel spreadsheet, whose function call in the 'main' section of the code can be commented in and out as needed.

*openpyxl, multiprocessing

*Remove comment symbol on the line 'generate_excel(transaction_id_start, product_id_range, price_range, num_sheets)' to generate a Test Excel file

*Can change variable Values on Main section, can comment out unwanted Functions

*Test .xlsx files ('transactions.xlsx', 'transactions_generated.xlsx') are provided in this Source Code

*Outputs on 'processed_files' folder

*Documentations on Doc folder

*Acknowledgements
1. Programming with Mosh- https://youtu.be/_uQrJ0TkZlc
