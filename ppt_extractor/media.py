from .utils import binary_to_decoded_bytes

def extract_media(shape):

    bytes_data = binary_to_decoded_bytes(shape.blob)
    content_type = shape.content_type
    ext = shape.ext
    file_name = shape.file_name
    sha1 = shape.sha1

    data = {
        "data" : bytes_data,
        "content_type" : content_type,
        "extension" : ext,
        "file_name" : file_name,
        "sha1_hash_digest" : sha1
    }

    return data