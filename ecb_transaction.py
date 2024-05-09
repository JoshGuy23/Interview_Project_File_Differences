# This class represents a row in the ecb file.
class Ecb:
    def __init__(self):
        # Initializes the class
        self.spa = ""
        self.service_code = ""
        self.charge = 0.0
        self.new_charge = 0.0

    def set_spa(self, new_spa):
        # The SPA of the row is converted to a string,
        # as some SPAs have alphabetical characters while others do not
        self.spa = str(new_spa)

    def set_service_code(self, new_code):
        self.service_code = new_code

    def set_charges(self, c, n_c):
        # Ensures that both the charge and new charge are floats
        # The dollar signs from the charge cells are removed
        self.charge = float(c[1:])
        self.new_charge = float(n_c)
