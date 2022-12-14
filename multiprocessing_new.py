import multiprocessing_new
import cProfile
import os
import pandas as pd
import parser_csv

class Statistic:
    def __init__(self, file: str, profession: str):
        """
        Инициализация объекта статистики
        Args:
            file: путь к файлу с данными для статистики
            profession: название професии
        """
        self.file = file
        self.profession = profession
        self.years_salary_dictionary = {}
        self.years_count_dictionary = {}
        self.years_salary_vacancy_dict = {}
        self.years_count_vacancy_dict = {}
        self.area_salary_dict = {}
        self.area_count_dict = {}

    def get_stat(self):
        self.get_stat_by_year_multi_on()
        self.get_stat_by_city()

    def get_stat_by_year(self, file_csv):
        """
        Составление статистики по каждому году
        Args:
            file_csv: название csv-файла
        """
        df = pd.read_csv(file_csv)
        df["salary"] = df[["salary_from", "salary_to"]].mean(axis=1)
        df["published_at"] = df["published_at"].apply(lambda s: int(s[:4]))
        df_vac = df[df["name"].str.contains(self.profession)]
        return df["published_at"].values[0], [int(df["salary"].mean()), len(df), int(df_vac["salary"].mean()), len(df_vac)]

    def get_stat_by_city(self):
        """
        Получение статистики вакансий по городам
        """
        df = pd.read_csv(self.file)
        total = len(df)
        df["salary"] = df[["salary_from", "salary_to"]].mean(axis=1)
        df["count"] = df.groupby("area_name")["area_name"].transform("count")
        df = df[df["count"] > total * 0.01]
        df = df.groupby("area_name", as_index=False)
        df = df[["salary", "count"]].mean().sort_values("salary", ascending=False)
        df["salary"] = df["salary"].apply(lambda s: int(s))
        self.area_salary_dict = dict(zip(df.head(10)["area_name"], df.head(10)["salary"]))
        df = df.sort_values("count", ascending=False)
        df["count"] = round(df["count"] / total, 4)
        self.area_count_dict = dict(zip(df.head(10)["area_name"], df.head(10)["count"]))

    def get_stat_by_year_multi_off(self):
        """
        Получение статистики по годам
        """
        res = []
        for csv_file in os.listdir("Csvs"):
            with open(os.path.join("Csvs", csv_file), "r") as file_csv:
                res.append(self.get_stat_by_year(file_csv.name))
        for year, data_stat in res:
            self.years_salary_dictionary[year] = data_stat[0]
            self.years_count_dictionary[year] = data_stat[1]
            self.years_salary_vacancy_dict[year] = data_stat[2]
            self.years_count_vacancy_dict[year] = data_stat[3]

    def get_stat_by_year_multi_on(self):
        """
        Собирает статистику по годам, с использованием мультипроцессорности
        """
        csv_file = [rf"Csvs\{file_name}" for file_name in os.listdir("Csvs")]
        pool = multiprocessing_new.Pool(4)
        res_list = pool.starmap(self.get_stat_by_year, [(file,) for file in csv_file])
        pool.close()

        for year, data_stat in res_list:
            self.years_salary_dictionary[year] = data_stat[0]
            self.years_count_dictionary[year] = data_stat[1]
            self.years_salary_vacancy_dict[year] = data_stat[2]
            self.years_count_vacancy_dict[year] = data_stat[3]

    def print_stat(self):
        print(f'Динамика уровня зарплат по годам: {self.years_salary_dictionary}')
        print(f'Динамика количества вакансий по годам: {self.years_count_dictionary}')
        print(f'Динамика уровня зарплат по годам для выбранной профессии: {self.years_salary_vacancy_dict}')
        print(f'Динамика количества вакансий по годам для выбранной профессии: {self.years_count_vacancy_dict}')
        print(f'Уровень зарплат по городам (в порядке убывания): {self.area_salary_dict}')
        print(f'Доля вакансий по городам (в порядке убывания): {self.area_count_dict}')

if __name__ == '__main__':
    file_path = "C:\\Users\\RBT\\PycharmProjects\\Korelina\\.idea\\Files\\vacancies_by_year.csv"
    prof = "Аналитик"
    parser_csv.parse_csv_by_year(file_path)
    stat = Statistic(file_path, prof)
    stat.get_stat()
    stat.print_stat()
    #cProfile.run("solve.get_stat_by_year_multi_on()", sort="cumtime")
    #cProfile.run("solve.get_stat_by_year_multi_off()", sort="cumtime")