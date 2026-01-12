import random
import math
from openpyxl.styles import Border, Side, Alignment


def random_range(min_val, max_val, floor=True):
    """生成随机数"""
    if floor:
        return random.randint(int(min_val), int(max_val))
    else:
        value = random.random() * (max_val - min_val) + min_val
        return round(value, 3)


def random_pick(arr, count):
    """从列表中随机选取指定数量的元素"""
    arr_copy = arr.copy()
    result = []
    for _ in range(count):
        if not arr_copy:
            break
        index = random.randint(0, len(arr_copy) - 1)
        item = arr_copy.pop(index)
        result.append(item)
    return result


def set_wrap_border(cell):
    """设置单元格边框和对齐方式"""
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    cell.border = thin_border
    cell.alignment = Alignment(vertical='center', horizontal='center')
