from pptx import Presentation
import base64

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
        # if shape.has_text_frame:
        #     print(shape.text)
        # Extracts data for each shape_type based on the switch condition
        slide_data.append(switch_type(shape.shape_type)(shape))

    return slide_data

def extract_shape(shape):
    return shape.shape_type

def extract_picture(shape):
    image_data = base64.b64encode(shape.image.blob).decode("utf-8")
    data = {
        "data" : image_data,
        "shape_type" : shape.shape_type
    }
    return data

def extract_text(shape):
    text_data = shape.text
    data = {
        "data" : text_data,
        "shape_type" : shape.shape_type
    }
    return data

def extract_placeholder(shape):
    if shape.has_text_frame:
        return extract_text(shape)
    data = {
        "data" : "Unknown Placeholder data",
        "shape_type" : shape.shape_type
    }
    return data

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
        14 : extract_placeholder, # "placeholder",
        16 : extract_shape, # "media",
        17 : extract_text, # "textbox",
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