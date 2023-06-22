from collections import defaultdict

from frames.model_frame import ModelFrame


class ZeroingFrame(ModelFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.name = 'Zeroing'
        self.building_diagram = {self.name: defaultdict(defaultdict)}
        self.floors_data = self.building_diagram[self.name]
        self.type_of_security = ['S301', 'S303', 'S191', 'S193', 'WT', 'BM']
        self.protection_type = ['B', 'A', 'C', 'D', 'F', 'gG']
        self.In = ['16', '2', '4', '6', '10', '20', '25', '32', '40', '50', '63', '80', '100', '125', '160', '315',
                   '400']
