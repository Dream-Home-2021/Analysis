# -*- coding: utf-8 -*-

import yagmail
import datetime as dt
import schedule


# 收件人
# to = ['cs4@12t.cn', 'huangtingting@xm12t.com', 'CS92@xm12t.com', 'cs93@xm12t.com', 'zhoulili@xm12t.com',
#       'lihua@xm12t.com', 'yexiaohuan@xm12t.com', 'jirongyu@xm12t.com', 'wuyanhong@xm12t.com', 'keyunping@xm12t.com']
# # 抄送
# to2 = ['cs@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com',
#        'luoxiucai@xm12t.com']
# to = ['fanglongsheng@xm12t.com']
# to2 = ['fanglongsheng@xm12t.com']
# filename = '新开部门行业户均监控-{}.xlsx'.format(yd.strftime('%m%d'))
#
# # 附件
# attach = '../新开户均/新开部门行业户均监控/' + filename
#
# # 标题
# subject = """
# 大家好：
#      附件为""" + str(yd.month) + """月新开部门行业户均完成日报，针对完成率较低的账户和部门，请尽快提升。谢谢
# 祝：
# 商祺
# """

# 个人信息


# def send_email():
#     yag.send(to, cc=to2, subject=filename, contents=[subject, html], attachments=[attach])
#     print(filename + "发送成功\n")


# schedule.every().day.at('23:00').do(send_email)


# while True:
#     schedule.run_pending()
#     time.sleep(1)


class SendEmailyao:
    def __init__(self, sj, cs, subject, contents, attachments, pretime):
        self.username = 'fanglongsheng@xm12t.com'
        self.password = '008759'
        self.yag = None
        self.sj = sj
        self.cs = cs
        self.subject = subject
        self.attachments = attachments
        self.pretime = pretime
        self.myIns = "./wo.html"
        self.contents = contents + [self.myIns]

    def send_email(self):
        self.yag = yagmail.SMTP(self.username, self.password)
        self.yag.send(to=self.sj, cc=self.cs, subject=self.subject, contents=self.contents,
                      attachments=self.attachments)

    def go(self):
        try:
            schedule.every().day.at(self.pretime).do(self.send_email())
            print(self.subject + "  发送成功\n")
        except Exception as e:
            if "the first argument must be callable" in str(e):
                print(self.subject + "  发送成功\n")
            else:
                print(str(e) + "\n" + self.subject + "  发送失败\n")
        pass


# # 大搜客户消费监控
yd = dt.datetime.today() - dt.timedelta(1)
#
# recipient = ['cs@12t.cn', 'cs4@12t.cn', 'cs11@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn',
#              'CS92@xm12t.com', 'luoxiucai@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com']
#
# cc = ['xiezq@xm12t.com', 'cs5@12t.cn', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com']
#
# body = """
# 大家好:
#    附件为""" + str(yd.month) + """月的各大区大搜客户消费，请查收！
# 祝:
# 商祺!
# """
#
# filename = '2023年大搜客户消费监控总表-{}.xlsx'.format(yd.strftime('%m%d'))
#
# attach = '../大搜客户消费监控/' + filename
#
# searchConsume = SendEmailyao(sj='fanglongsheng@xm12t.com', cs='fanglongsheng@xm12t.com', subject=filename, contents=[body],
#                  attachments=attach, pretime='10:47')
# searchConsume.go()
#
# # 信息流客户消费监控
# recipient2 = ['cs@12t.cn', 'cs4@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn', 'CS92@xm12t.com', 'luoxiucai@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com', 'cs11@12t.cn']
# cc2 = ['cs5@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs1@12t.cn']
# body2 = """
# 大家好:
#    附件为""" + str(yd.month) + """月的各大区信息流消费，请查收！
# 祝:
# 商祺!
# """
# filename2 = '2023年信息流客户消费监控总表-{}.xlsx'.format(yd.strftime('%m%d'))
# attach2 = '../信息流客户消费监控/' + filename2
#
# feedConsume = SendEmailyao(sj='fanglongsheng@xm12t.com', cs='fanglongsheng@xm12t.com', subject=filename2, contents=[body2],
#                  attachments=attach2, pretime='10:47')
# feedConsume.go()
#
#
#
# # 未消费客户明细
#
# # 计算当前季度
# current_quarter = (dt.datetime.today().month - 1) // 3 + 1
#
# recipient3 = ['cs4@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn', 'CS92@xm12t.com', 'luoxiucai@xm12t.com', 'cs93@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'lihua@xm12t.com', 'yexiaohuan@xm12t.com', 'chenyuan@xm12t.com', 'jirongyu@xm12t.com', 'cs11@12t.cn']
# cc3 = ['cs20@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs1@12t.cn']
# body3 = """
# 大家好：
#       附件为2022年Q"""+ str(current_quarter) +"""截止昨日开户未消费的客户明细，请各主管协助组员填写未消费原因，谢谢！
# 祝：
# 商祺！
# """
# filename3 = '2023Q'+str(current_quarter)+'未消费客户明细-{}.xlsx'.format(yd.strftime('%m%d'))
# attach3 = '../未消费客户明细/账户未消费' + filename3
# noConsume1 = SendEmailyao(sj='fanglongsheng@xm12t.com', cs='fanglongsheng@xm12t.com', subject=filename3, contents=[body3],
#                  attachments=attach3, pretime='10:47')
# noConsume1.go()
#
#
#
# # 未消费客户明细-销售
# # 计算当前季度
# current_quarter = (dt.datetime.today().month - 1) // 3 + 1
# recipient4 = ['shuwg@12t.cn', 'linyd@12t.cn', 'lindepei@xm12t.com', 'wengquanlong@xm12t.com', 'wangdefa@xm12t.com']
# cc4 = ['wangdefa@xm12t.com', 'cs1@12t.cn']
# body4 = """
# 大家好：
#       附件为2023年Q"""+ str(int(current_quarter)-1) +"""截止昨日开户未消费的客户明细，请各主管协助组员填写未消费原因，谢谢！
# 祝：
# 商祺！
# """
# filename4 = '2023Q'+str(current_quarter)+'未消费客户明细-{}.xlsx'.format(yd.strftime('%m%d'))
# attach4 = '../未消费客户明细/发给销售部门的账户未消费' + filename4
# noConsume2 = SendEmailyao(sj='fanglongsheng@xm12t.com', cs='fanglongsheng@xm12t.com', subject=filename4, contents=[body4],
#                  attachments=attach4, pretime='10:47')
# noConsume2.go()
#
#
#
#
# # 新开户均
# recipient5 = ['cs4@12t.cn', 'huangtingting@xm12t.com', 'CS92@xm12t.com', 'cs93@xm12t.com', 'zhoulili@xm12t.com',
#       'lihua@xm12t.com', 'yexiaohuan@xm12t.com', 'jirongyu@xm12t.com', 'wuyanhong@xm12t.com', 'keyunping@xm12t.com']
# cc5 = ['cs@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com',
#        'luoxiucai@xm12t.com']
# filename5 = '新开部门行业户均监控-{}.xlsx'.format(yd.strftime('%m%d'))
# attach5 = '../新开户均/新开部门行业户均监控/' + filename5
# body5 = """
# 大家好：
#      附件为""" + str(yd.month) + """月新开部门行业户均完成日报，针对完成率较低的账户和部门，请尽快提升。谢谢
# 祝：
# 商祺
# """
# newAc = SendEmailyao(sj='fanglongsheng@xm12t.com', cs='fanglongsheng@xm12t.com', subject=filename5, contents=[body5],
#                  attachments=attach5, pretime='10:47')
# newAc.go()
#
# #大搜客户失效监控
# recipient6 = ['cs@12t.cn', 'cs4@12t.cn', 'cs11@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn', 'CS92@xm12t.com']
# cc6 = ['cs5@12t.cn', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs93@xm12t.com', 'yeling@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com', 'lihua@xm12t.com', 'yexiaohuan@xm12t.com', 'jiangmei@xm12t.com', 'cs24@12t.cn', 'chenyuan@xm12t.com', 'chenqiaoyun@xm12t.com', 'zengdawei@xm12t.com', 'jirongyu@xm12t.com', 'duxinxin@xm12t.com', 'wuyanhong@xm12t.com', 'youruifeng@xm12t.com']
#
# filename6 = '2023年搜索客户失效监控-{}.xlsx'.format(yd.strftime('%m%d'))
# attach6 = '../失效监控/大搜客户失效监控/' + filename6
# body6 = """
# 大家好:
#        附件为""" + str(yd.month) + """月份搜索失效客户，如有续费请在续费栏填写已续费，如未续费请注明未续费原因，
# 客户失效3天内还未续费的客户，请主管协助组员分析失效原因并实施挽救
# 祝:
#        商祺！
# """
# failureDetSearch = SendEmailyao(sj='fanglongsheng@xm12t.com', cs='fanglongsheng@xm12t.com', subject=filename6, contents=[body6],
#                  attachments=attach6, pretime='10:47')
# failureDetSearch.go()
#
#
# #大搜客户失效监控
# recipient7 = ['xiaoyuemei@xm12t.com', 'zhoulili@xm12t.com']
# cc7 = ['cs5@12t.cn', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs93@xm12t.com', 'yeling@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com', 'lihua@xm12t.com', 'yexiaohuan@xm12t.com', 'jiangmei@xm12t.com', 'cs24@12t.cn', 'chenyuan@xm12t.com', 'chenqiaoyun@xm12t.com', 'zengdawei@xm12t.com', 'jirongyu@xm12t.com', 'duxinxin@xm12t.com', 'wuyanhong@xm12t.com', 'youruifeng@xm12t.com']
#
# filename7 = '2023年信息客户失效监控-{}.xlsx'.format(yd.strftime('%m%d'))
# attach7 = '../失效监控/信息流客户失效监控/' + filename7
# body7 = """
# 大家好:
#        附件为""" + str(yd.month) + """月份搜索失效客户，如有续费请在续费栏填写已续费，如未续费请注明未续费原因，
# 客户失效3天内还未续费的客户，请主管协助组员分析失效原因并实施挽救
# 祝:
#        商祺！
# """
# failureDetfeed = SendEmailyao(sj='fanglongsheng@xm12t.com', cs='fanglongsheng@xm12t.com', subject=filename7, contents=[body7],
#                  attachments=attach7, pretime='10:47')
# failureDetfeed.go()

#周报-框架
recipient8 = ['xm@xm12t.com']
filename8 = '框架客户代理商信息周报-厦门易尔通 -{}.xlsx'.format(yd.strftime('%m%d'))
filename8_1 = '框架客户统计-华南区(第37周数据).xlsx'
attach8 = '../../每周/框架客户代理商信息周报（周1的9点半前--数据源）/' + filename8
attach8_1 = '../../每周/框架客户代理商信息周报（周1的9点半前--数据源）/' + filename8_1
body8 = """
附件为框架客户代理商信息周报，请查收，谢谢！
"""
failureDetfeed = SendEmailyao(sj=recipient8, cs=None ,subject=filename8, contents=[body8],
                 attachments=[attach8,attach8_1], pretime='10:47')
failureDetfeed.go()