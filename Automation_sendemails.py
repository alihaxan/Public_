import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime


def send_email(subject, body, to_email):
    # Replace the placeholders below with your email and password or use environment variables.
    from_email = 'your_email@gmail.com'
    email_password = 'your_email_password'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))  # Set the email content type to HTML

    # Set up the SMTP server and send the email
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp_server:
        smtp_server.starttls()
        smtp_server.login(from_email, email_password)
        smtp_server.send_message(msg)


def main():
    # Pick up the required Excel file from your computer
    excel_file_path = 'path_to_your_excel_file.xlsx'

    # List of sheet numbers for clients you want to send reminders to
    sheet_numbers = [1, 2, 3]  # Replace with your desired sheet numbers

    # Load the Excel file
    sheets_data = pd.read_excel(excel_file_path, sheet_name=sheet_numbers, engine='openpyxl')

    # Create a DataFrame to store the results
    results_df = pd.DataFrame(columns=['Timestamp', 'Name', 'Email', 'Message', 'Status'])

    # Get the current date and time
    current_datetime = datetime.now()
    current_datetime_str = current_datetime.strftime("%Y-%m-%d %H:%M:%S")

    # Define the filename for the results Excel file
    results_filename = f"results_{current_datetime.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"

    for sheet_number, data in sheets_data.items():
        for index, row in data.iterrows():
            client_email = row['Email']
            start_row = 5  # Replace with the starting row number for the table
            end_row = 10  # Replace with the ending row number for the table
            start_column = 2  # Replace with the starting column number for the table
            end_column = 7  # Replace with the ending column number for the table

            # Extract the table from the DataFrame based on the specified row and column numbers
            table_data = data.iloc[start_row:end_row + 1, start_column:end_column + 1]

            # Get the name of the client from the 'Name' column
            client_name = row['Name']

            # Create a personalized message with the name and table
            message = f"Dear {client_name},<br><br>"
            message += "Here are your client details:<br>"
            message += table_data.to_html(index=False)

            # Send the email to the client
            send_email('Monthly Payment Reminder', message, client_email)
            print(f"Reminder sent to {client_email}.")

            # Update the results DataFrame with the sent email details and status
            results_df = results_df.append({
                'Timestamp': current_datetime_str,
                'Name': client_name,
                'Email': client_email,
                'Message': message,
                'Status': 'done'
            }, ignore_index=True)

    # Save the results DataFrame to the specified path
    results_df.to_excel(results_filename, index=False, engine='openpyxl')
    print(f"Results saved to {results_filename}.")


if __name__ == "__main__":
    main()
