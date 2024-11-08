from pptx import Presentation
from .slide import extract_slide

# Extracts the data from the PPT file
def load(file):
    ppt = Presentation(file)
    ppt_data = {}
    for idx, slide in enumerate(ppt.slides):
        ppt_data[idx] = extract_slide(slide)
    return ppt_data