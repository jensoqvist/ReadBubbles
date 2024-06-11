from pos_number import PositionNumber

class PositionNumbers():
    def __init__(self, pos_list_in) -> None:
        self.pos_list_in = pos_list_in
        self.position_numbers = []
        self.set_position_numbers()
        self.lenght = len(self.position_numbers[0])

    def set_position_numbers(self):
        for pos in self.pos_list_in:
            self.position_numbers.append(PositionNumber(pos).list_values())
        self.position_numbers = sorted(self.position_numbers)


if __name__ == "__main__":
    pass