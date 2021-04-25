import math
import openpyxl as xl
import pathlib


def P(m, j, ρ: list):  # Вероятность нахождения аппаратной системы m в сосоянии j
    def product_1():  # Числитель
        result = 1
        for i in range(len(m)):
            result *= math.factorial(m[i]) / (math.factorial(m[i] - j[i]) * math.factorial(j[i])) * (ρ[i] ** j[i])
        return result

    def product_2():  # Знаменатель
        result = 1
        for i in range(len(m)):
            result *= (ρ[i] + 1) ** m[i]
        return result

    return product_1() / product_2()


def G(m, ρ, M):  # Надёжность аппаратной системы
    N = len(m)

    def recursive_sum(j, i=0):
        if i == N:
            return P(m, j, ρ)
        result = 0
        for j[i] in range(M[i], m[i]+1):
            result += recursive_sum(j, i+1)
        return result

    j = [0] * N  # Состояние аппаратной системы (сколько исправно)
    return recursive_sum(j[:])  # Передать копию списка, чтобы не модифицировать исходный


if __name__ == '__main__':
    # Открыть файл Excel
    filepath = pathlib.Path(__file__).parent / 'data.xlsx'
    wb = xl.load_workbook(filepath)
    ws = wb['Входные данные']

    # Загрузить значения из файла
    ν = [x.value for x in list(ws.rows)[1]][1:]  # Параметр интенсивности
    μ = [x.value for x in list(ws.rows)[2]][1:]  # Параметр обслуживания
    ρ = [ν / μ for ν, μ in zip(ν, μ)]
    m = [x.value for x in list(ws.rows)[3]][1:]  # Конфигурация аппаратной система (сколько всего)
    M = [x.value for x in list(ws.rows)[4]][1:]  # Минимальное требуемое для обеспечения заданной производительности число исправных процессоров

    print('Входные данные:', ν, μ, ρ, m, M, sep='\n')
    print(f'Надежность {G(m, ρ, M)}')
    # input('Enter - выйти')