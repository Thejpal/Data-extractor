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
        
        # Switches between extractors for each shape based on shape type
        else:
            slide_data.append(switch_type(shape.shape_type)(shape))

    return slide_data

# Switches to different data extractors based on the shape type. Returns the function object which can be run in the original function
def switch_type(shape_type):
    switch = {
        -2 : "mixed",
        1 : "autoshape",
        2 : "callout",
        3 : "chart",
        4 : "comment",
        5 : "freeform",
        6 : "extract_data",
        7 : "embedded_ole_object",
        8 : "form_control",
        9 : "line",
        10 : "linked_ole_object",
        11 : "linked_picture",
        12 : "ole_control_object",
        13 : "picture",
        14 : "placeholder",
        16 : "media",
        17 : "textbox",
        18 : "script_anchor",
        19 : "table",
        20 : "canvas",
        21 : "diagram",
        22 : "ink",
        23 : "ink_comment",
        24 : "igx_graphic",
        25 : "",
        26 : "web_video"
    }
    return switch.get(shape_type)