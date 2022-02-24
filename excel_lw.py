# -*- coding: utf-8 -*-
"""
## Лабораторная работа
"""

import numpy as np
import pandas as pd
import xlwings as xw

"""1. Загрузите данные из файлов `reviews_sample.csv` (__ЛР2__) и `recipes_sample_with_tags_ingredients.csv` (__ЛР5__) в виде `pd.DataFrame`. Обратите внимание на корректное считывание столбца(ов) с индексами. Оставьте в таблице с рецептами следующие столбцы: `id`, `name`, `minutes`, `submitted`, `description`, `n_ingredients`"""

rs = pd.read_csv('reviews_sample.csv')
rswti = pd.read_csv('recipes_sample_with_tags_ingredients.csv')
rs

rswti = rswti[['id', 'name', 'minutes', 'submitted', 'description', 'n_ingredients']]
rswti

"""2. Случайным образом выберите 5% строк из каждой таблицы и сохраните две таблицы на разные листы в один файл `recipes.xlsx`. Дайте листам названия "Рецепты" и "Отзывы", соответствующие содержанию таблиц. """

rs5=rs.sample(frac=0.05)
rswti5=rswti.sample(frac=0.05)
rswti5

writer = pd.ExcelWriter('recipes.xlsx', engine='xlsxwriter')

rswti5.to_excel(writer, sheet_name='Рецепты')
rs5.to_excel(writer, sheet_name='Отзывы')
writer.save()

"""3. Используя `xlwings`, добавьте на лист `Рецепты` столбец `seconds_assign`, показывающий время выполнения рецепта в секундах. Выполните задание при помощи присваивания массива значений диапазону ячеек."""

secas = rswti5['minutes']*60

dt = xw.Book('recipes.xlsx')
sh=dt.sheets['Рецепты']
sh.range('H1').value = secas

"""4. Используя `xlwings`, добавьте на лист `Рецепты` столбец `seconds_formula`, показывающий время выполнения рецепта в секундах. Выполните задание при помощи формул Excel."""

sh.range('J1').value = 'seconds_formula'
sh.range('J2:J1501').formula = f'=(D2:D1501)*60'

"""5. Добавьте на лист `Рецепты`  столбец `n_reviews`, содержащий кол-во отзывов для этого рецепта. Выполните задание при помощи формул Excel.

6. Сделайте названия всех добавленных столбцов полужирными и выровняйте по центру ячейки.
"""

sh.range(f'$A1:$J1').api.HorizontalAlignment = -4108
sh.range(f'$A1:$J1').api.Font.Bold = True

"""7. Раскрасьте ячейки столбца `minutes` в соответствии со следующим правилом: если рецепт выполняется быстрее 5 минут, то цвет - зеленый; от 5 до 10 минут - жёлтый; и больше 10 - красный."""

for cl in sh.range('D2:D1501'):
    if cl.value<5: cl.color = (0,128,0)
    elif cl.value>10: cl.color = (255,0,0)
    else: cl.color = (255,255,0)

"""8. Напишите функцию `validate()`, которая проверяет соответствие всех строк из листа `Отзывы` следующим правилам:
    * Рейтинг - это число от 0 до 5 включительно
    * Соответствующий рецепт имеется на листе `Рецепты`
    
В случае несоответствия этим правилам, выделите строку красным цветом
"""

sh2=dt.sheets['Отзывы']
for cl in sh2.range('F2:F1501'):
    if cl.value>5: cl.color = (255,0,0)