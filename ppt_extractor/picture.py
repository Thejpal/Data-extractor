from .utils import binary_to_decoded_bytes

def extract_picture(shape):
    image_data = binary_to_decoded_bytes(shape.image.blob)
    data = {
        "data" : image_data,
        "shape_type" : shape.shape_type
    }
    return data