class NumberNotInRange(Exception):
    def __init__(self):
        self.message = "Number is not in range (0-100)"
        super().__init__(self.message)


class EmptyFields(Exception):
    def __init__(self):
        self.message = "The fields should not be empty"
        super().__init__(self.message)
