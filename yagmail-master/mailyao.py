# -*- coding: utf-8 -*-

import yagmail
import datetime as dt
import schedule
import time

# 大搜
dasou = 0
# 信息流
xxliu = 0

# 季度表
jidu = 0
# 未消费-客服
weixiaofei_kefu = 0
# 未消费-销售
weixiaofei_xiaos = 0
# 新开
xingkai = 0
# 大搜失效
sixiao_dasou = 0
# 信息流失效
sixiao_xxliu = 0
# 框架周报
kuangjia = 0
# 布瑞泽翼百信更近
ybx_brz = 0
# 布瑞泽
ybx = 0
# 翼百信
brz = 0

class SendEmailyao:
    def __init__(self, sj, cs, subject, contents, attachments, pretime):
        self.username = 'fanglongsheng@xm12t.com'
        self.password = 'fls520ly.'
        self.yag = yagmail.SMTP(self.username, self.password)
        self.sj = sj
        self.cs = cs
        self.subject = subject
        self.attachments = attachments
        self.pretime = pretime
        self.myIns = "./wo.html"
        self.contents = contents + [self.myIns]
        self.task_completed = False

    def send_email(self):
        # self.yag = yagmail.SMTP(self.username, self.password)
        self.yag.send(to=self.sj, cc=self.cs, subject=self.subject, contents=self.contents,
                      attachments=self.attachments)
        self.task_completed = True

    def go(self):
        try:
            schedule.every().day.at(self.pretime).do(self.send_email)
            while not self.task_completed:
                # 获取当前时间
                current_time = dt.datetime.now()

                # 提取分钟和秒数
                minutes = current_time.minute
                seconds = current_time.second
                print(minutes, seconds)

                schedule.run_pending()
                time.sleep(1)

            print(self.subject + "  发送成功\n")
        except Exception as e:
            if "the first argument must be callable" in str(e):
                print(self.subject + "  发送成功\n")
            else:
                print(str(e) + "\n" + self.subject + "  发送失败\n")
        pass



# # 大搜客户消费监控
yd = dt.datetime.today() - dt.timedelta(1)
if dasou:
    recipient = ['cs@12t.cn', 'cs4@12t.cn', 'cs11@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn',
                 'CS92@xm12t.com', 'luoxiucai@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com']

    cc = ['xiezq@xm12t.com', 'cs5@12t.cn', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com',
          'fanglongsheng@xm12t.com']

    body = """
    大家好:
       附件为""" + str(yd.month) + """月的各大区大搜客户消费，请查收！
    祝:
    商祺!
    """

    filename = '2023年大搜客户消费监控总表-{}.xlsx'.format(yd.strftime('%m%d'))

    attach = '../大搜客户消费监控/' + filename

    searchConsume = SendEmailyao(sj=recipient, cs=cc, subject=filename,
                                 contents=[body],
                                 attachments=attach, pretime='08:58')
    searchConsume.go()

# 信息流客户消费监控
if xxliu:
    recipient2 = ['cs@12t.cn', 'cs4@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn', 'CS92@xm12t.com',
                  'luoxiucai@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com', 'cs11@12t.cn']
    cc2 = ['cs5@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs1@12t.cn', 'fanglongsheng@xm12t.com']
    body2 = """
    大家好:
       附件为""" + str(yd.month) + """月的各大区信息流消费，请查收！
    祝:
    商祺!
    """
    filename2 = '2023年信息流客户消费监控总表-{}.xlsx'.format(yd.strftime('%m%d'))
    attach2 = '../信息流客户消费监控/' + filename2

    feedConsume = SendEmailyao(sj=recipient2, cs=cc2, subject=filename2,
                               contents=[body2],
                               attachments=attach2, pretime='09:02')
    feedConsume.go()

# # 未消费客户明细
if weixiaofei_kefu:
    # 计算当前季度
    current_quarter = (dt.datetime.today().month - 1) // 3 + 1

    recipient3 = ['cs4@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn', 'CS92@xm12t.com', 'luoxiucai@xm12t.com',
                  'cs93@xm12t.com', 'cs12@12t.cn', 'yecaihua@xm12t.com', 'lihua@xm12t.com', 'yexiaohuan@xm12t.com',
                  'chenyuan@xm12t.com', 'jirongyu@xm12t.com', 'cs11@12t.cn']
    cc3 = ['cs20@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs1@12t.cn', 'fanglongsheng@xm12t.com']
    body3 = """
    大家好：
          附件为2022年Q""" + str(current_quarter) + """截止昨日开户未消费的客户明细，请各主管协助组员填写未消费原因，谢谢！
    祝：
    商祺！
    """
    filename3 = '2023Q' + str(current_quarter) + '未消费客户明细-{}.xlsx'.format(yd.strftime('%m%d'))
    attach3 = '../未消费客户明细/账户未消费/' + filename3
    noConsume1 = SendEmailyao(sj=recipient3, cs=cc3, subject=filename3,
                              contents=[body3],
                              attachments=attach3, pretime='09:44')
    noConsume1.go()

# # 未消费客户明细-销售
if weixiaofei_xiaos:
    # 计算当前季度
    current_quarter = (dt.datetime.today().month - 1) // 3 + 1
    recipient4 = ['shuwg@12t.cn', 'linyd@12t.cn', 'lindepei@xm12t.com', 'wengquanlong@xm12t.com', 'wangdefa@xm12t.com']
    cc4 = ['wangdefa@xm12t.com', 'cs1@12t.cn', 'fanglongsheng@xm12t.com']
    body4 = """
    大家好：
          附件为2023年Q""" + str(int(current_quarter) - 1) + """截止昨日开户未消费的客户明细，请各销售部门主管协助组员关注和跟进，谢谢！
    祝：
    商祺！
    """
    filename4 = '2023Q' + str(current_quarter) + '未消费客户明细-{}.xlsx'.format(yd.strftime('%m%d'))
    attach4 = '../未消费客户明细/发给销售部门的账户未消费/' + filename4
    noConsume2 = SendEmailyao(sj=recipient4, cs=cc4, subject=filename4, contents=[body4],
                              attachments=attach4, pretime='09:51')
    noConsume2.go()
# 新开
if xingkai:
    # 收件人
    to = ['cs4@12t.cn', 'huangtingting@xm12t.com', 'CS92@xm12t.com', 'cs93@xm12t.com', 'zhoulili@xm12t.com',
          'lihua@xm12t.com', 'yexiaohuan@xm12t.com', 'jirongyu@xm12t.com', 'wuyanhong@xm12t.com', 'keyunping@xm12t.com']
    # 抄送
    to2 = ['cs@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com',
           'luoxiucai@xm12t.com', 'fanglongsheng@xm12t.com']
    # to = ['fanglongsheng@xm12t.com']
    # to2 = ['fanglongsheng@xm12t.com']
    filename = '新开部门行业户均监控-{}.xlsx'.format(yd.strftime('%m%d'))

    # 附件
    attach = '../新开户均/新开部门行业户均监控/' + filename

    # 标题
    subject = """
    大家好：
         附件为""" + str(yd.month) + """月新开部门行业户均完成日报，针对完成率较低的账户和部门，请尽快提升。谢谢
    祝：
    商祺
    """
    newkai = SendEmailyao(sj=to, cs=to2, subject=filename, contents=[subject],
                          attachments=[attach], pretime='10:05')
    newkai.go()

# #大搜客户失效监控
if sixiao_dasou:
    recipient6 = ['cs@12t.cn', 'cs4@12t.cn', 'cs11@12t.cn', 'cs20@12t.cn', 'huangtingting@xm12t.com', 'cs33@12t.cn',
                  'CS92@xm12t.com']
    cc6 = ['cs5@12t.cn', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs93@xm12t.com', 'yeling@xm12t.com',
           'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com', 'lihua@xm12t.com', 'yexiaohuan@xm12t.com',
           'jiangmei@xm12t.com', 'cs24@12t.cn', 'chenyuan@xm12t.com', 'chenqiaoyun@xm12t.com', 'zengdawei@xm12t.com',
           'jirongyu@xm12t.com', 'duxinxin@xm12t.com', 'wuyanhong@xm12t.com', 'youruifeng@xm12t.com',
           'fanglongsheng@xm12t.com']

    filename6 = '2023年搜索客户失效监控-{}.xlsx'.format(yd.strftime('%m%d'))
    attach6 = '../失效监控/大搜客户失效监控/' + filename6
    body6 = """
    大家好:
           附件为""" + str(yd.month) + """月份搜索失效客户，如有续费请在续费栏填写已续费，如未续费请注明未续费原因，
    客户失效3天内还未续费的客户，请主管协助组员分析失效原因并实施挽救
    祝:
           商祺！
    """
    failureDetSearch = SendEmailyao(sj=recipient6, cs=cc6, subject=filename6, contents=[body6],
                                    attachments=attach6, pretime='10:17')
    failureDetSearch.go()

# #信息流客户失效监控
if sixiao_xxliu:
    recipient7 = ['xiaoyuemei@xm12t.com', 'zhoulili@xm12t.com']
    cc7 = ['cs5@12t.cn', 'cs1@12t.cn', 'xugzh@12t.cn', 'wangdefa@xm12t.com', 'cs93@xm12t.com', 'yeling@xm12t.com',
           'cs12@12t.cn', 'yecaihua@xm12t.com', 'cs88@xm12t.com', 'lihua@xm12t.com', 'yexiaohuan@xm12t.com',
           'jiangmei@xm12t.com', 'cs24@12t.cn', 'chenyuan@xm12t.com', 'chenqiaoyun@xm12t.com', 'zengdawei@xm12t.com',
           'jirongyu@xm12t.com', 'duxinxin@xm12t.com', 'wuyanhong@xm12t.com', 'youruifeng@xm12t.com',
           'fanglongsheng@xm12t.com']

    filename7 = '2023年信息客户失效监控-{}.xlsx'.format(yd.strftime('%m%d'))
    attach7 = '../失效监控/信息流客户失效监控/' + filename7
    body7 = """
    大家好:
           附件为""" + str(yd.month) + """月份搜索失效客户，如有续费请在续费栏填写已续费，如未续费请注明未续费原因，
    客户失效3天内还未续费的客户，请主管协助组员分析失效原因并实施挽救
    祝:
           商祺！
    """
    failureDetfeed = SendEmailyao(sj=recipient7, cs=cc7, subject=filename7,
                                  contents=[body7],
                                  attachments=attach7, pretime='11:41')
    failureDetfeed.go()

# 周报-框架
if kuangjia:
    recipient8 = ['xm@xm12t.com']
    ccx = ['fanglongsheng@xm12t.com']
    filename8 = '框架客户代理商信息周报-厦门易尔通 -{}.xlsx'.format(yd.strftime('%m%d'))
    filename8_1 = '框架客户统计-华南区(第37周数据).xlsx'
    attach8 = '../../每周/框架客户代理商信息周报（周1的9点半前--数据源）/' + filename8
    attach8_1 = '../../每周/框架客户代理商信息周报（周1的9点半前--数据源）/' + filename8_1
    body8 = """
    附件为框架客户代理商信息周报，请查收，谢谢！
    """
    failureDetfeed = SendEmailyao(sj=recipient8, cs=ccx, subject=filename8, contents=[body8],
                                  attachments=[attach8, attach8_1], pretime='09:29')
    failureDetfeed.go()



# 季度表
if jidu:
    recipient10 = ['xugzh@12t.cn']
    cc10 = ['cs@12t.cn', 'cs4@12t.cn', 'cs20@12t.cn', 'cs1@12t.cn', 'wangdefa@xm12t.com', 'fanglongsheng@xm12t.com']


    filename10 = '2023年Q4季度任务监控总表-{}.xlsx'.format(yd.strftime('%m%d'))
    attach10= '../季度任务监控/' + filename10
    body10= """
    大家好：
        附件是2023年Q4季度任务监控总表数据，请查收，谢谢！
    祝：
    商祺！
    """
    jidu = SendEmailyao(sj=recipient10, cs=cc10, subject=filename10,
                                  contents=[body10],
                                  attachments=attach10, pretime='14:30')
    jidu.go()

# # 翼百信布瑞泽
# if ybx_brz:
#     recipient11 = ['xugzh@12t.cn']
#     cc11 = ['cs@12t.cn', 'cs4@12t.cn', 'cs20@12t.cn', 'cs1@12t.cn', 'wangdefa@xm12t.com', 'fanglongsheng@xm12t.com']
#
#     filename11 = '2023年Q4季度任务监控总表-{}.xlsx'.format(yd.strftime('%m%d'))
#     attach11 = '../季度任务监控/' + filename11
#     body11 = """
#     大家好：
#         附件是2023年Q4季度任务监控总表数据，请查收，谢谢！
#     祝：
#     商祺！
#     """
#     jidu = SendEmailyao(sj=recipient11, cs=cc11, subject=filename11,
#                         contents=[body11],
#                         attachments=attach11, pretime='14:30')
#     jidu.go()