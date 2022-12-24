import main2
import main4

input_data = input('Вакансии или Статистика: ')
if __name__ == '__main__':
    if input_data == 'Вакансии':
        main2.main()
    elif input_data == 'Статистика':
        main4.main()