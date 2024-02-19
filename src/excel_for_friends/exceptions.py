class NumberNotInRange(Exception):
    def __init__(self):
        self.message = "Number is not in range (0-100)"
        super().__init__(self.message)

