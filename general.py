#! /usr/bin/env python

import messytables
import xypath

from os.path import join as pjoin

filename = pjoin('fixtures', 'basic-multidim-test.xls')
messy = messytables.excel.XLSTableSet(open(filename, "rb"))
table = xypath.Table.from_messy(messy.tables[0])
#print xypath.Bag(table=table) in table
#raise SystemExit

#print table.filter("years").x

def f(table):
    
    cell_category = table.filter("category").assert_one()

    categories = cell_category.fill(xypath.DOWN).nonblank()

    years = list(table.filter("years").assert_one().fill(xypath.RIGHT).nonblank())


    def months(year, next_year):
        if next_year is None:
            return year.shift(x=-1, y=1).fill(xypath.RIGHT)
        return list(table.range(year, next_year).shift(y=1))[:-1]

    def value():
        return table.get_at(month.x, status.y).value

    for category in categories:
        product = category.shift(x=1)
        units_list = product.shift(1, 1).extrude(0, 2)

        for units in units_list:
            status = units.shift(1)

            for year, next_year in zip(years, years[1:] + [None]):
                for month in months(year, next_year):
                    yield (
                        category.value,
                        product.value,
                        units.value,
                        status.value,
                        year.value,
                        month.value,
                        value())

def f(table):

    categories = column(...)
    years = row(...)
    units_list = (...)
    ...

    for <stuff> in g(categories, years, units_list):
        

for category, product, unit, status, year, month, value in f(table, extraction_description):
    print category, product, unit, status, year, month, value
