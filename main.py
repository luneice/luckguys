import hashlib
import numpy
from xlrd import open_workbook
from xlwt import Workbook


# 20201116 期双色球开奖结果
obj = "05061416192710"
obj_hash = hashlib.sha1(obj.encode("utf-8")).hexdigest()


# 计算两个 HASH 字符串的欧式距离
# 把 HASH 字符串看成字符向量，字符从 16 进制转成 10 进制，然后计算欧式距离
def distance(target, obj):
	sum = 0
	for i, j in zip(target, obj):
		# 单个字符转 10 进制，然后算欧式距离
		sum += numpy.square(int(i, 16) - int(j, 16))
		pass
	# 返回距离
	return numpy.sqrt(sum)


def write_excel(file, sheet_name, col_name_list, col_part_list, col_phone_list, col_hash_list, col_dis_list):

	sheet1 = file.add_sheet(sheet_name, cell_overwrite_ok=True)
	row_title = ["姓名", "学院", "手机号", "手机号对应的 HASH 值（SHA1）", "距离"]

	# 写第一行
	for i in range(0, len(row_title)):
		sheet1.write(0, i, row_title[i])
		pass

	# 写第一列
	for i in range(0, len(col_name_list)):
		sheet1.write(i + 1, 0, col_name_list[i])
		pass

	# 写第二列
	for i in range(0, len(col_part_list)):
		sheet1.write(i + 1, 1, col_part_list[i])
		pass

	# 写第三列
	for i in range(0, len(col_phone_list)):

		sheet1.write(i + 1, 2, col_phone_list[i])
		pass

	# 写第四列
	for i in range(0, len(col_hash_list)):
		sheet1.write(i + 1, 3, col_hash_list[i])
		pass

	# 写第五列
	for i in range(0, len(col_dis_list)):
		sheet1.write(i + 1, 4, col_dis_list[i])
		pass
	pass


# 读取 excel 表格的信息
with open_workbook('namelist.xlsx') as workbook:

	# 用于输出结果的表格
	write_sheet = Workbook()

	# 遍历表格
	for sheet in workbook.sheet_names():
		worksheet = workbook.sheet_by_name(sheet)
		col_name_list = []
		col_part_list = []
		col_phone_list = []
		col_hash_list = []
		col_dis_list = []
		# 按行读取信息
		for row_index in range(worksheet.nrows):
			if row_index == 0:
				continue
				pass
			name = worksheet.cell(row_index, 0).value
			part = worksheet.cell(row_index, 1).value

			phone = str(worksheet.cell(row_index, 2).value)[:11]
			phone_hash = hashlib.sha1(phone.encode("utf-8")).hexdigest()
			# 脱敏处理
			phone_item = phone[:3] + "****" + phone[7:]
			dis = distance(phone_hash, obj_hash)

			col_name_list.append(name)
			col_part_list.append(part)
			col_phone_list.append(phone_item)
			col_hash_list.append(phone_hash)
			col_dis_list.append(dis)
			pass
		write_excel(write_sheet, sheet, col_name_list, col_part_list, col_phone_list, col_hash_list, col_dis_list)

	write_sheet.save("lucky_guys.xlsx")
	pass
