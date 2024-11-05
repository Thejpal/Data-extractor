from pptx import Presentation

# Extracts the data from the PPT file
def load(file):
    ppt = Presentation(file)
    ppt_data = {}
    for idx, slide in enumerate(ppt.slides):
        ppt_data[idx] = extract_data(slide)
    return ppt_data

# Extracts data from a PPT slide
def extract_data(slide):
    slide_data = []

    for shape in slide.shapes:
        # Recursively extracts data for shape_type "group"
        if shape.shape_type == 6:
            data = {"group" : extract_data(shape)}
            slide_data.append(data)
        # Add code for other shape types

    return slide_data