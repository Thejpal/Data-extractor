def extract_text(shape):
    text_data = shape.text
    data = {
        "data" : text_data,
        "shape_type" : shape.shape_type
    }
    return data