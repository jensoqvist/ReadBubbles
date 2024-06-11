import json

class Settings():
    def __init__(self) -> None:
        self.file = "settings.json"
        self.data = self.get_data()

    def get_data(self):
        with open(self.file) as json_file:
            data = json.load(json_file)
        return data

if __name__ == '__main__':
    pass

