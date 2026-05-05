from pos_number import PositionNumber
from settings import Settings

class PositionNumbers():
    def __init__(self, pos_list_in, settings= Settings().data) -> None:
        self.pos_list_in = pos_list_in
        self.position_numbers = []
        self.settings= settings
        self.set_position_numbers()
        self.lenght = len(self.position_numbers[0])

    def set_position_numbers(self):
        for pos in self.pos_list_in:
            self.position_numbers.append(PositionNumber(pos, settings= self.settings).list_values())
        self.position_numbers = sorted(self.position_numbers)


if __name__ == "__main__":
    pass