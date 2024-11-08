from .table import extract_table
from .textbox import extract_text
from .shape import extract_shape

def extract_placeholder(shape):
    if shape.has_text_frame:
        return extract_text(shape)
    elif shape.has_table:
        return extract_table(shape)
    elif shape.has_chart:
        return extract_shape(shape)
    data = {
        "data" : "Unknown Placeholder data",
        "shape_type" : shape.shape_type
    }
    return data