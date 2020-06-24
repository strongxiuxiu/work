import xlrd  # 引入模块
import psycopg2



def xlrd_inster(n,num):
	conn_1 = psycopg2.connect(database="x", user="postgres", password="x", host="x", port="5432")
	cur1 = conn_1.cursor()
	# 通过sql中join的方式查出未爬到的数据
	# 打开文件，获取excel文件的workbook（工作簿）对象
	workbook = xlrd.open_workbook("/home/pxq/test.xls")  # 文件路径
	class_body = workbook.sheet_by_name(sheet_name="GJB832A小类")
	item_list = []
	all= []
	for i in range(num):
		# print(i)
		# print(i)
		# all.append(i)
		try:
			mdir = class_body.row_values(i + 1)
			# print(mdir[0])
			if i < n:
				itme = class_body.row_values(i + 1)
				if itme[0] == "序号":
					pass
				else:
					pass
					item_list.append(itme[2])

					# sql1 = 'INSERT INTO "public"."small_Classify"("small_classname","belong_classcode")VALUES'+"('%s','832A')"%itme[2]
					# # sql1 = 'INSERT INTO "small_Classify" (small_classname,belong_classcode) VALUES (%s,"832A")'
					# # db = cur1.execute(sql1, [itme[2]])
					# print(sql1)
					# cur1.execute(sql1)
					# conn_1.commit()
				# INSERT INTO "public"."small_Classify"("ID", "small_classname", "belong_classcode") VALUES (1, '1', '1');
			else:
				itme = class_body.row_values(i + 1)

				if itme[2] in item_list:
					num = i + 1
					while True:

						it = class_body.row_values(num + 1)
						# print(v.row_values(num))
						all.append(class_body.row_values(num))
						if it[2] in item_list:
							name = it[2]
							global name
							print(name,33333)
							break
						else:

							num += 1
							print(it)
							all.append(it)
							sql1 = 'INSERT INTO "public"."small_Catalog"("Catalog_classname","belong_Catalog","smallcode")VALUES' + "('%s','%s','%s')" % \
								   (it[2],name,it[1])
							print(sql1)
							cur1.execute(sql1)
							conn_1.commit()
		except:
			break

