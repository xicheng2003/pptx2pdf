import os
import comtypes.client

def ppt_to_pdf(input_file, output_file):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(input_file)
    presentation.SaveAs(output_file, 32)  # 32 represents the PDF format
    presentation.Close()
    powerpoint.Quit()

def batch_convert_ppt_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(".ppt") or filename.endswith(".pptx"):
            input_file = os.path.join(input_folder, filename)
            output_file = os.path.join(output_folder, filename.rsplit('.', 1)[0] + ".pdf")
            ppt_to_pdf(input_file, output_file)
            print(f"Converted {filename} to PDF.")

input_folder = "your_file_path"
output_folder = "your_file_path"

batch_convert_ppt_to_pdf(input_folder, output_folder)