import pandas as pd


class TimeStatistics:
    def __init__(self):
        self.origin_data = {}
        self.department_data = {}

    def __call__(self, *args, **kwargs):
        pass

    def data_init(self):
        day_num = [str(i) for i in range(32)]
        pd_data: dict = pd.read_csv('time.csv').to_dict()
        for i in range(len(pd_data['姓名'])):
            department_key = pd_data['部门'][i]
            if department_key not in self.department_data:
                self.department_data[department_key] = []
            department_data = self.department_data[department_key]
            for key, data in pd_data.items():
                if key not in day_num:
                    continue
                time_str = str(data[i]).strip()
                if '\n' in time_str:
                    time_str = time_str.split('\n')[1]
                department_data.append(time_str)
        for key in self.department_data:
            print(self.department_data[key])

    def individual2department(self):
        pass


if __name__ == '__main__':
    ts = TimeStatistics()
    ts.data_init()
