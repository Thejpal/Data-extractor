from .shape import extract_shape
from .picture import extract_picture
from .table import extract_table
from .placeholder import extract_placeholder
from .textbox import extract_text

# Extracts data from a PPT slide
def extract_slide(slide):
    slide_data = []

    for shape in slide.shapes:
        if shape.shape_type == 3:
            print("chart")
        # Extracts data for each shape_type based on the switch condition
        slide_data.append(switch_type(shape.shape_type)(shape))

    return slide_data

# Switches to different data extractors based on the shape type. Returns the function object which can be run in the original function
def switch_type(shape_type):
    switch = {
        -2 : extract_shape, # "mixed",
        1 : extract_shape, # "autoshape",
        2 : extract_shape, # "callout",
        3 : extract_shape, # "chart",
        4 : extract_shape, # "comment",
        5 : extract_shape, # "freeform",
        6 : extract_slide, # "group"
        7 : extract_shape, # "embedded_ole_object",
        8 : extract_shape, # "form_control",
        9 : extract_shape, # "line",
        10 : extract_shape, # "linked_ole_object",
        11 : extract_shape, # "linked_picture",
        12 : extract_shape, # "ole_control_object",
        13 : extract_picture, # "picture",
        14 : extract_placeholder, # "placeholder",
        16 : extract_shape, # "media",
        17 : extract_text, # "textbox",
        18 : extract_shape, # "script_anchor",
        19 : extract_table, # "table",
        20 : extract_shape, # "canvas",
        21 : extract_shape, # "diagram",
        22 : extract_shape, # "ink",
        23 : extract_shape, # "ink_comment",
        24 : extract_shape, # "igx_graphic",
        25 : extract_shape, # "",
        26 : extract_shape, # "web_video"
    }
    return switch.get(shape_type)