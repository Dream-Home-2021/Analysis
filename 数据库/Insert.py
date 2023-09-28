import pymysql
import openpyxl

wb = openpyxl.load_workbook('./XX连锁超市数据/商品销售明细表.xlsx')
sheet = wb['未命名组件']

# 手动输入
# no = int(input('部门编号: '))
# name = input('部门名称: ')
# location = input('部门所在地: ')

# 1. 创建连接（Connection）
conn = pymysql.connect(host='localhost', port=3306,
                       user='root', password='fls520ly',
                       database='analysis_test', charset='utf8mb4')
cursor = conn.cursor()
for i, row in enumerate(sheet.iter_rows(min_row=2), 1):
    data = [cell.value for cell in row]
    cursor.execute('insert into sales values(%s, %s, %s, %s, %s, %s, %s, %s)',
                   data)
conn.commit()
#
#
# try:
#     # 2. 获取游标对象（Cursor）
#     with conn.cursor() as cursor:
#         # 3. 通过游标对象向数据库服务器发出SQL语句
#         affected_rows = cursor.execute(
#             'insert into `store_data` values (%s, %s, %s)'
#             # ,(no, name, location)
#         )
#         if affected_rows == 1:
#             print('新增部门成功!!!')
#     # 4. 提交事务（transaction）
#     conn.commit()
# except pymysql.MySQLError as err:
#     # 4. 回滚事务
#     conn.rollback()
#     print(type(err), err)
# finally:
#     # 5. 关闭连接释放资源
#     conn.close()
