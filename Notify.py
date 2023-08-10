import requests

class Notify():
    def __init__(self, message, image_path):
        headers = {
            'Authorization':'Bearer ' + '2WBWz8va4c0iNIADsw4pNpDTyBF9skWKxWNl8D6EJDm'
        }

        data = { 'message': message}
        image = open(image_path, 'rb')
        files = { 'imageFile': image}

        requests.post('https://notify-api.line.me/api/notify',headers = headers, data = data, files = files)