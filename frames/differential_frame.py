from collections import defaultdict

from frames.model_frame import ModelFrame


class DifferentialFrame(ModelFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.name = 'Differential'
        self.building_diagram = {self.name: defaultdict(defaultdict)}
        self.floors_data = self.building_diagram[self.name]
        self.type_of_security = ['P302', 'P304']
        self.protection_type = ['AC', 'A', 'AC-s', 'A-s']
        self.In = ['0.03', '0.006', '0.01', '0.1']
