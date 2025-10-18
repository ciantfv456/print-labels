from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import barcode
from barcode.writer import ImageWriter
import csv
from pydantic import BaseModel, ConfigDict, field_validator
import pandas as pd
import win32api
from win32printing import win32print


class Label(BaseModel):
    product: str
    kasher: str
    barcode: str
    month_offset: int
    model_config = ConfigDict(extra='ignore')
    amount: int

    @field_validator('barcode', "kasher", "product", mode='before')
    @classmethod
    def convert_int_to_str(cls, v):
        if isinstance(v, int):
            return f"0{str(v)}"
        return v[::-1]

    @property
    def today(self):
        return f"{datetime.today().date().day}\{datetime.today().date().month}\{datetime.today().date().year}" + "תאריך ייצור: "[::-1]

    @property
    def due(self):
        due_output = int(datetime.today().date().month + self.month_offset) % 12 if int(datetime.today().date().month + self.month_offset) % 12 != 0 else 12
        return f"{datetime.today().date().day}\{due_output}\{datetime.today().date().year}" + "לשימוש עד: "[::-1]

    @property
    def back_text(self):
        return ''.join(["רכיבים: "[::-1], "100%", " בשר עגל מובחר"[::-1] ][::-1]) +  "\nהוראות אחסנה: לשמור בקירור\n"[::-1]  + ' '.join(["של"[::-1] ,  "2-0"  ,   "מעלות עד השימוש"[::-1]][::-1])

def from_excel_to_code(path):
    df = pd.read_excel(path)
    
    # Get the column names
    columns = df.columns
    
    # Create a list to hold instances of ExcelRecord
    records = []

    # Iterate over each row in the DataFrame
    for _, row in df.iterrows():
        # Convert the row to a dictionary
        row_dict = row.to_dict()
        # Create an instance of ExcelRecord passing the row data
        record = Label(**row_dict)
        records.append(record)

    return records

# TODO: when printing a lot of labels the order of the printed labels is getting mixed
def print_label(path):
    printers=win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
    PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}
    temprint=printers[1][2]
    handle = win32print.OpenPrinter(temprint, PRINTER_DEFAULTS)
    level = 2
    attributes = win32print.GetPrinter(handle, level)
    attributes['pDevMode'].PaperWidth = 5000
    attributes['pDevMode'].PaperLength = 10000
    attributes['pDevMode'].PaperSize =0
    print(win32print.SetPrinter(handle, level, attributes, 0))
    win32api.ShellExecute(0,'printto',path,'"%s"' % temprint,'.',0)
    win32print.ClosePrinter(handle)

def convert_cm_to_pixels(cm, dpi):
    """Converts a centimeter value to pixels based on the specified DPI."""
    return int((cm / 3) * dpi)

def create_barcode(data):
    # Define the barcode type and data
    barcode_class = barcode.get_barcode_class('Code128')

    # Create the barcode object with ImageWriter
    my_barcode = barcode_class(data, writer=ImageWriter())

    # Save the barcode as an image file
    return my_barcode.save('my_ean13_barcode')

def create_templated_image_cm(
    template_path,
    logo_path,
    barcode_path,
    font_path,
    output_path,
    title_text,
    today_text,
    due_text,
    back_text,
    kashrut,
    kashrut_text,
    width_cm,
    height_cm,
    dpi=300
):
    try:
        # Calculate pixel dimensions based on desired cm and DPI
        width_px = convert_cm_to_pixels(width_cm, dpi)
        height_px = convert_cm_to_pixels(height_cm, dpi)

        # Open and resize the template image to the new pixel dimensions
        img = Image.open(template_path).resize((width_px, height_px)).convert("RGBA")
        draw = ImageDraw.Draw(img)

        # Add a logo to the image
        logo = Image.open(logo_path).convert("RGBA")
        logo_size = (int(width_px / 7), int(height_px / 7))
        logo = logo.resize(logo_size)
        
        # Define the position for the logo (e.g., top-right corner)
        logo_position = (convert_cm_to_pixels(2, dpi)), int(convert_cm_to_pixels(4, dpi))
        img.paste(logo, logo_position, logo)

        # Add a logo to the image
        barcode = Image.open(barcode_path).convert("RGBA")
        barcode_size = (int(width_px / 3), int(height_px / 7))
        barcode = barcode.resize(barcode_size)
        
        # Define the position for the logo (e.g., top-right corner)
        barcode_position = (convert_cm_to_pixels(3.5, dpi)), int(convert_cm_to_pixels(6.2, dpi))
        img.paste(barcode, barcode_position, barcode)

        # Load fonts
        # You may need to adjust font sizes relative to the image's new dimensions
        title_font = ImageFont.truetype(font_path, size=int(width_px / 16))
        date_font = ImageFont.truetype(font_path, size=int(width_px / 30))
        kashrut_font = ImageFont.truetype(font_path, size=int(width_px / 30))
        back_font = ImageFont.truetype(font_path, size=int(width_px / 30))

        # Add text to the image
        draw.text((width_px / 2, convert_cm_to_pixels(3.7, dpi)), title_text, font=title_font, fill=(0, 0, 0), anchor="mm", align="center")

        draw.text((convert_cm_to_pixels(9, dpi), convert_cm_to_pixels(5.5, dpi)), today_text, font=date_font, fill=(0, 0, 0), anchor="rt", align="center")
        draw.text((convert_cm_to_pixels(9, dpi), convert_cm_to_pixels(5.8, dpi)), due_text, font=date_font, fill=(0, 0, 0), anchor="rt", align="center")

        draw.multiline_text((convert_cm_to_pixels(5, dpi), convert_cm_to_pixels(4.1, dpi)), back_text, align="right",  font=back_font, fill=(50, 50, 50))
        draw.multiline_text((convert_cm_to_pixels(1.5, dpi), convert_cm_to_pixels(5.4, dpi)), kashrut + "\n" + kashrut_text, font=kashrut_font, fill=(0, 0, 0), align="center")


        # Save the final image with the correct DPI metadata
        img.save(output_path, dpi=(dpi, dpi))
        img.resize((width_px * 1.5, height_px * 1.5), Image.LANCZOS)
        img.save(output_path, dpi=(dpi, dpi))
        print(f"Image saved to {output_path}")

    except FileNotFoundError as e:
        print(f"Error: A required file was not found. Please check your file paths. {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


# --- Usage Example ---
if __name__ == "__main__":
    # The actual function call with your paths, text, and desired dimensions in cm
    labels = from_excel_to_code("מדבקות.xlsx")
    for l in labels:
        if l.amount > 0:
            create_templated_image_cm(
                template_path=r"C:\Users\shama\OneDrive\Desktop\meat.jpg",
                logo_path=r"C:\Users\shama\Documents\מדבקות\assets\000.jpg",
                barcode_path=create_barcode(l.barcode),
                font_path="arial.ttf", # You must use a valid font file
                output_path=f"{l.barcode}.png",
                title_text=l.product,
                today_text=l.today,
                due_text=l.due,
                back_text=l.back_text,
                kashrut=l.kasher,
                kashrut_text="בהשגחת רבנות כרמיאל"[::-1],
                width_cm=10,
                height_cm=10,
                dpi=300 # Set to 96 for web images, 300+ for print
            )
            for _ in range(l.amount):
                print_label(f"{l.barcode}.png")

