class Ecb:
    def __init__(self):
        self.spa = ""
        self.service_code = ""
        self.charge = 0.0
        self.new_charge = 0.0

    def set_spa(self, new_spa):
        self.spa = str(new_spa)

    def set_service_code(self, new_code):
        self.service_code = new_code

    def set_charges(self, c, n_c):
        self.charge = float(c[1:])
        self.new_charge = float(n_c)
