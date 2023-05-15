import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from tkinter import Tk, Button
from tkinter.filedialog import askopenfilename
import fitz
import datetime 
import win32com.client as win32
import datetime
import os
from pdf2image import convert_from_path
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
# print(desktop_path)

max_width = 0
total_height = 0



import pathlib

# Get the path to the desktop directory
desktop_path2 = pathlib.Path.home() / "Desktop"

# Define the directory name
directory_name1 = "Output"
# Define the directory name
directory_name2 = "Stacked"

# Create the full path to the directory
directory_path1 = os.path.join(desktop_path2, directory_name1)
directory_path2 = os.path.join(desktop_path2, directory_name2)

# Check if the directory already exists
if not os.path.exists(directory_path1):
    # Create the directory if it doesn't exist
    os.makedirs(directory_path1)
# Check if the directory already exists
if not os.path.exists(directory_path2):
    # Create the directory if it doesn't exist
    os.makedirs(directory_path2)



stacked_photos_dir=desktop_path+"\\Stacked\\"
output_folder =desktop_path+"\\Output"
# stacked_photos_dir = 'C:\\Users\\vb\\Desktop\\tomail\\'
# output_folder = "C:\\Users\\vb\\Desktop\\Output Dashboard"
# powerbi_dir = "C:\\Users\\vb\\AppData\\Local\\Temp\\Power BI Desktop"
powerbi_dir=desktop_path



def get_file_path():
    home_dir = os.path.expanduser(stacked_photos_dir)
    file_pathz = filedialog.askopenfilename(initialdir=home_dir)
    print(f"The selected file path is: {file_pathz}")
    return file_pathz


# Create the main GUI window
window = tk.Tk()
window.title("DB-Generator")
window.geometry("500x250")
window.configure(bg="#FFFFFF")
# Set a consistent color scheme
primary_color = "#2c3e50"
secondary_color = "#FFFFFF"


window.iconbitmap('1.ico')


pdf_file_path = ""

def select_pdf_file():
    # Ask user to select a PDF file
    global pdf_file_path
    pdf_file_path = askopenfilename(initialdir=powerbi_dir,filetypes=[("PDF Files", "*.pdf")])
    if pdf_file_path:
        print(f"Selected PDF file: {pdf_file_path}")
        # Perform further processing on the PDF file
# print(out)
# To TIFF
def convert_pdf_to_images(pdf_path, output_folder):
    with fitz.open(pdf_path) as pdf_doc:
        num_pages = pdf_doc.page_count
    #Change dpi up to 600 for Higher quailty
    images = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=num_pages,poppler_path='C:\\poppler-utils\\bin')
    # images = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=num_pages)
 
    for i in range(num_pages):
        images[i].save(f"{output_folder}\\page_{i+1}.tiff", 'TIFF')
    
def convert_and_select():
    select_pdf_file()
    if pdf_file_path:
        convert_pdf_to_images(pdf_file_path, output_folder)

def select_photos():
    home_dir2 = os.path.expanduser(output_folder)
    # file_paths = filedialog.askopenfilenames(title="Select Photos", filetypes=(("PNG files", "*.png"),("JPEG files", "*.jpg")),initialdir=home_dir2)
    file_paths = filedialog.askopenfilenames(title="Select Photos",initialdir=home_dir2)

    photo_stack(file_paths)

def photo_stack(file_paths):
    global max_width
    global total_height
    images = []
    max_width = 0
    total_height = 0

    for file_path in file_paths:
        image = Image.open(file_path)
        max_width = max(max_width, image.width)
        total_height += image.height 
        images.append(image)

    stacked_image = Image.new("RGB", (max_width, total_height), (255, 255, 255))
    y_offset = 0

    for image in images:
        stacked_image.paste(image, (0, y_offset))
        y_offset += image.height
        # y_offset = total_height  # set y_offset to the current total 
    # Add datetime to filename
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"stacked_photo_{current_datetime}.png"
    stacked_image.save(stacked_photos_dir+filename)
    stacked_image.show()

def send_to_mail(mail_to,mail_Subject,mail_img):  

    # # get the current date
    # today = datetime.date.today()
    # # get the week number with ISO format
    # week_number = today.isocalendar()[1]
    # # format the week number as "WK" followed by the week number
    # week_number_str = f"WK{week_number:02d}"
    # create an instance of the Outlook application
    outlook = win32.Dispatch('outlook.application')
    # create a new email message
    mail = outlook.CreateItem(0)

    # set the recipient email address

    mail.To = mail_to
    # print(max_width)
    # print(total_height)

    # print(week_number_str)  # prints "WK19" (for example, if today is in the 19th week of the year)
    # set the subject of the email
    
    # mail.Subject = mail_Subject+week_number_str


    # get the HTML body of the email
    html_body = mail.HTMLBody

    # add the mail body text to the email body
    mail_body_text = "<br><br>"
    html_body = html_body + f'<p>{mail_body_text}</p>'

    # add the image to the email body
    image_path = mail_img
    cid = 'image1'
    html_body = html_body + f'<img src="cid:{cid}" width="99%">'
    mail.HTMLBody = html_body

    # attach the image to the email
    attachment = mail.Attachments.Add(Source=image_path, Type=5, DisplayName='image.jpg')

    # set the Content-ID of the image attachment
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)

    # display the email message
    mail.Display(True)

def send_to_mail_button():
    # list of dashboard options
    # mail_Subject = selected_option.get()
    mail_img = get_file_path()  
    # mail_to = mail_to_entry.get()
    # get_file_path()
    send_to_mail(" "," ",mail_img)


pdf_to_img_button = tk.Button(window, text="Select PDF", command=convert_and_select, bg=primary_color, fg=secondary_color, font=("Helvetica", 12),width=15)
pdf_to_img_button.pack(pady=15)


# Create the "Select Photos" button
select_button = tk.Button(window, text="Select Photos", command=select_photos, bg=primary_color, fg=secondary_color, font=("Helvetica", 12),width=15)
select_button.pack(pady=20)

# options = ['WE Air Weekly ','Indigo 150+ Weekly ','Network Weekly Dashboard ']
# # variable to save the value
# selected_option = tk.StringVar()
# # set the default value for the variable
# selected_option.set(options[0])

# # create the drop-down menu
# drop_down = tk.OptionMenu(window, selected_option, *options)
# drop_down.pack(pady=20)


# mail_to_entry = tk.Entry(window)
# mail_to_label = tk.Label(window, text='Mail to :')
# mail_to_entry.pack()
# mail_to_label.pack()


# button1 = tk.Button(window, text='Select Dashboard to send', command=get_file_path)
# button1.pack(pady=20)


# Send mail button
send_mail_button = tk.Button(window, text="Send Mail", command=send_to_mail_button, bg=primary_color, fg=secondary_color, font=("Helvetica", 12),width=15)
send_mail_button.pack(pady=20)


# Run the GUI
window.mainloop()