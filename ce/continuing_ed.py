import os
import pandas as pd
from pptx import Presentation, shapes
import win32com.client


## Functions:
def replace_powerpoint_w_pdf(powerpoint_path: str, pdf_path: str) -> None:
    """
    """
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    deck = powerpoint.Presentations.Open(powerpoint_path)
    deck.SaveAs(pdf_path, 32) 
    deck.Close()
    os.remove(powerpoint_path)
    return

def update_save_powerpoint(textbox: shapes.autoshape.Shape, updated_text: str, attendee_ce_pptx_path: str) -> None:
    """
    """
    paragraph = textbox.text_frame.paragraphs[0]
    paragraph.runs[0].text = updated_text
    ce_template.save(attendee_ce_pptx_path)
    return


## Global variables:
ROOT_DIRECTORY = r"C:\Users\sputnam\Documents\AWRA\11_7_22_Seminar"
CE_DIRECTORY = os.path.join(ROOT_DIRECTORY, "CE")
CE_FILENAME = "AWRA_CEC_11_7_22_Seminar_{}.pptx"


## Read in the attendance list:
attendees_filename = "november_attendees.xlsx"
attendees_path = os.path.join(ROOT_DIRECTORY, attendees_filename)

attendees = pd.read_excel(attendees_path)
attendees.columns = attendees.columns.str.lower()
attendees = attendees[attendees["attended"]].copy()

print(f"There were {attendees.shape[0]} and they paid ${attendees['amount'].sum()}\n")


## Read in the CE template:
ce_template_filename = "awra_ncrs_ce_template.pptx"
ce_template_path = os.path.join(ROOT_DIRECTORY, ce_template_filename)

ce_template = Presentation(ce_template_path)
ce_template_slide = ce_template.slides[0]


## For eacha attendee make a certificate:
if not os.path.exists(CE_DIRECTORY):
    os.makedirs(CE_DIRECTORY)

textbox = [shape for shape in ce_template_slide.shapes if shape.shape_type == 17 and 'First Last' in shape.text][0]

for i, attendee in enumerate(attendees.itertuples()):
    first_name = attendee.first
    last_name = attendee.last

    attendee_ce_pptx_path = os.path.join(CE_DIRECTORY, CE_FILENAME.format(last_name))
    attendee_ce_pdf_path = attendee_ce_pptx_path.replace(".pptx", ".pdf")
    updated_text = f"{first_name} {last_name}"

    update_save_powerpoint(textbox, updated_text, attendee_ce_pptx_path)

    replace_powerpoint_w_pdf(attendee_ce_pptx_path, attendee_ce_pdf_path)
    
    if i+1 == attendees.shape[0]:
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Quit()

