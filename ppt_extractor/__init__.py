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
        # Extracts data for each shape_type based on the switch condition
        slide_data.append(switch_type(shape.shape_type)(shape))

    return slide_data

def extract_shape(shape):
    return shape.shape_id

# Switches to different data extractors based on the shape type. Returns the function object which can be run in the original function
def switch_type(shape_type):
    switch = {
        -2 : extract_shape, # "mixed",
        1 : extract_shape, # "autoshape",
        2 : extract_shape, # "callout",
        3 : extract_shape, # "chart",
        4 : extract_shape, # "comment",
        5 : extract_shape, # "freeform",
        6 : extract_data, # "group"
        7 : extract_shape, # "embedded_ole_object",
        8 : extract_shape, # "form_control",
        9 : extract_shape, # "line",
        10 : extract_shape, # "linked_ole_object",
        11 : extract_shape, # "linked_picture",
        12 : extract_shape, # "ole_control_object",
        13 : extract_shape, # "picture",
        14 : extract_shape, # "placeholder",
        16 : extract_shape, # "media",
        17 : extract_shape, # "textbox",
        18 : extract_shape, # "script_anchor",
        19 : extract_shape, # "table",
        20 : extract_shape, # "canvas",
        21 : extract_shape, # "diagram",
        22 : extract_shape, # "ink",
        23 : extract_shape, # "ink_comment",
        24 : extract_shape, # "igx_graphic",
        25 : extract_shape, # "",
        26 : extract_shape, # "web_video"
    }
    return switch.get(shape_type)