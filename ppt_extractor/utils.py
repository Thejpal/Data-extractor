import base64

def binary_to_decoded_bytes(bytes):
    return base64.b64encode(bytes).decode("utf-8")