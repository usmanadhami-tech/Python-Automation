# Python-Automation
Connecting with MS outlook and attaching the excel file to be sent via email.
#to interact with outlook
import win32com.client as win32

#defining parameters
def send_email_with_attachment(to, subject, body, attachment_path):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(to)  # Join multiple recipients with semicolon
    mail.Subject = subject
    mail.Body = body
    attachment = mail.Attachments.Add(attachment_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "MyAttachment")
    mail.Send()

#entry point of the function
def main():
    # Email details
    to = ['mohammad.usman@zooplus.com', 'a@example.com', 'b@example.com']
    subject = 'Price data attached'
    body = 'Please find the attached Excel file.'

    # Path to the Excel file
    excel_file_path = 'C:\\Users\\mohammad.adhami\\Downloads\\python_scripts\\price_1000_data.xlsx'

    # Send email with attachment
    send_email_with_attachment(to, subject, body, excel_file_path)

#function is executed only if the script is run directly (not imported as a module into another script)
if __name__ == "__main__":
    main()
