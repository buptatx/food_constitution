#! -*- coding:utf-8 -*-


import time
import os
import xlrd
import xlwt


class FoodConstitution(object):
    def __int__(self):
        self.constitution_foods = dict()

    @staticmethod
    def load_foods(filename):
        if not os.path.exists(filename):
            print("file %s not exists" % filename)
            return {}

        cf = dict()
        with open(filename, "r") as mf:
            for line in mf:
                temp = line.strip().split(":")
                if len(temp) < 2:
                    continue
                constitution = temp[0]
                foods = temp[1].split("、")
                if constitution in cf:
                    cf[constitution].extend(foods)
                else:
                    cf[constitution] = foods

        return cf

    @staticmethod
    def load_foods_from_excel(filename):
        if not os.path.exists(filename):
            print("file %s not exists" % filename)
            return {}

        foods = dict()
        wb = xlrd.open_workbook(filename)
        sh = wb.sheet_by_index(0)
        for i in range(1, sh.nrows):
            food_name = sh.cell_value(i, 3)
            food_id = sh.cell_value(i, 2)
            if food_name in foods:
                print("%s is duplicated" % food_name)
                continue

            foods[food_name] = food_id

        return foods

    def get_constitution_food_id(self, food_file, constitution_food_file):
        foods = FoodConstitution.load_foods_from_excel(food_file)
        constitution_foods = FoodConstitution.load_foods(constitution_food_file)

        if len(foods) < 1 or len(constitution_foods) < 1:
            print("invalid input")
            return -1

        res = list()
        not_find_foods = list()
        for constitution in constitution_foods:
            for food in set(constitution_foods[constitution]):
                if food in foods:
                    res.append([constitution, food, foods[food]])
                else:
                    res.append([constitution, food, "None"])
                    not_find_foods.append(food)
                    print("can not find %s in %s" % (food, food_file))

        for item in set(not_find_foods):
            print(item)

        self.store_res(res)

    def store_res(self, result):
        if len(result) == 0:
            print("no data for store")
            return -1

        res_file_name = time.strftime("./result/res.%Y%m%d.%H%M%S.xlsx")
        wb = xlwt.Workbook(encoding="utf-8")
        sh = wb.add_sheet("体质食物表")

        sh.write(0, 0, "体质")
        sh.write(0, 1, "食物名称")
        sh.write(0, 2, "食物ID")

        idx = 1
        for item in result:
            sh.write(idx, 0, item[0])
            sh.write(idx, 1, item[1])
            sh.write(idx, 2, item[2])
            idx += 1

        wb.save(res_file_name)


if __name__ == "__main__":
    test = FoodConstitution()
    test.get_constitution_food_id("./data/tbl_food.xlsx", "./data/体质食物.txt")