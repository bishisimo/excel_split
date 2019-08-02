import pandas as pd
import plotly.offline as py
import plotly.graph_objs as go


class TimeStatistics:
    def __init__(self):
        self.origin_data = {}
        self.department_data = {}
        self.time_max = 0

    def __call__(self, *args, **kwargs):
        self.data_init()
        self.data2department()
        # self.vis()

    def data_init(self):
        day_num = [str(i) for i in range(32)]
        pd_data: dict = pd.read_csv('time.csv').to_dict()
        for i in range(len(pd_data['姓名'])):
            department_key = pd_data['部门'][i]
            if department_key not in self.origin_data:
                self.origin_data[department_key] = []
            department_data = self.origin_data[department_key]
            for key, data in pd_data.items():
                if key not in day_num:
                    continue
                time_str = str(data[i]).strip()
                if '\n' in time_str:
                    time_str = time_str.split('\n')[1]
                if '出差' in time_str:
                    print('jia')
                    time_str = '19:30'
                time_str = time_str.split(':')

                if len(time_str) == 2:
                    sum_min = (int(time_str[0]) - 17) * 60 + int(time_str[1]) - 30
                    result = sum_min // 60
                    if result < 0:
                        continue
                    if result > self.time_max:
                        self.time_max = result
                    department_data.append(result)
        # for key in self.origin_data:
        #     print(key, self.origin_data[key])
        # print(self.time_max)

    def data2department(self):
        for key in self.origin_data:
            self.department_data[key] = [0] * (self.time_max + 1)
            department_data = self.department_data[key]
            for v in self.origin_data[key]:
                department_data[v] += 1
        # for key in self.department_data:
        #     print(key, self.department_data[key])

    def vis(self):
        x = [i for i in range(6)]
        show_data = [self.department_data[key] for key in self.department_data]
        # trace1=go.Bar(
        #     x=,
        #     y=show_data[0]
        # )
        # trace1 = [go.Bar(
        #     x=[i for i in range(6)],
        #     y=show_data[0]
        # )
        data = [go.Bar(x=x, y=self.department_data[key], name=key) for key in self.department_data]

        py.plot(data)


if __name__ == '__main__':
    ts = TimeStatistics()()
