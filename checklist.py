#!/usr/bin/python
# -*- coding: UTF-8 -*-
"""
20201109 hhw
更新add_sum,添加了通过率，失败率等

20190914 lizhao
update to VCS version

20190819 lizhao
fix bug name_check kill excel progress

20190816 lizhao
1 modify sumlist
2 add step state check
3 add step overall state check

20190805 lizhao
1 initial release
"""
import xlwings as xw
import os
import sys
import time
import logging
import pandas as pd

# 获取xlsx文件名
def get_xlsx():
    a = []
    path = os.getcwd()
    for x in os.listdir(path):
        if x.endswith('xlsx') and not x.startswith('~$') and not x.endswith('Release_Notes.xlsx'):
            a.append(x)
    return a


# 数字转字母
def get_char(number):
    factor, moder = divmod(number, 26)
    modchar = chr(moder + 65)
    if factor != 0:
        modchar = get_char(factor-1) + modchar
    return modchar


def write_log(wbname):
    global logger
    logger = logging.getLogger(wbname)
    logger.setLevel(logging.ERROR)

    fh = logging.FileHandler(f'{wbname}{time.strftime("%Y-%m-%d-%H-%M-%S")}.log')
    fh.setLevel(logging.ERROR)

    ch = logging.StreamHandler()
    ch.setLevel(logging.ERROR)

    formatter = logging.Formatter('[%(asctime)s][%(levelname)s] ## %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)

    logger.addHandler(fh)
    logger.addHandler(ch)


def init():
    global app
    app=xw.App(visible=True,add_book=False)
    app.display_alerts=False
    app.screen_updating=False


def teardown():
    app=xw.App(visible=True)
    app.display_alerts=True
    app.screen_updating=True

# read data
def read_data(bookname):
    global wb, cv, tc, tcdata, sumdata
    allname = []
    tmpname = ['1.Cover_Changelog', 'TestCase']
    wb = app.books.open(bookname)
    logger.critical(f'0.开始检查表格:\t{bookname}')
    for sheet in wb.sheets:
        allname.append(sheet.name)
    if not set(tmpname).issubset(set(allname)):
        logger.error(f'0.工作表名称错误:\t请使用{tmpname}')
        wb.save()
        time.sleep(1)
        wb.close()
        time.sleep(1)
        app.quit()
        time.sleep(1)
        sys.exit()

    cv = wb.sheets[0]
    tc = wb.sheets['TestCase']
    tcdata = pd.read_excel(bookname,sheet_name='TestCase')
    # 去除空值
    sumdata1 = tcdata[['Test Case Name', 'Test Case Description', 'Test Case RequirementID','Result Overall State', 'Result Details']]
    sumdata = sumdata1.dropna(subset=['Test Case Name', 'Test Case Description', 'Test Case RequirementID', 'Result Overall State'])


def save_quit():
    wb.save()
    time.sleep(1)
    wb.close()
    time.sleep(1)
    app.quit()


# 检查空值显示行号
def check_none(column, column_name):
    # print(column , '开始')
    rows_with_test_nane = []
    rows_with_step = []
    rows_difference = []
    for index_name, row_name in tcdata['Test Case Name'].items():
        if not pd.isna(row_name):
            rows_with_test_nane.append(index_name+1)
    for index_no, row in tcdata[column_name].items():
        if pd.isna(row):
            rows_with_step.append(index_no+2)
            rows_difference = set(rows_with_step).difference(set(rows_with_test_nane))
    if rows_difference is not None:
        # print(rows_difference, '输出')
        logger.error(f'首先检查TestCaseName是否为空!!!!如果CaseName存在的场合，请看左边->>>>> 项目 #{column_name}#: {column}列:的第{rows_difference}->该行有可能存在问题快去检查吧！')
    time.sleep(1)


# 检查总数不一样显示行号
def check_count(column, column_name):
    rows_with_nane = []
    rows_with_nane1 = []
    rows_with_step = []
    rows_with_step1 = []
    rows_difference = []
    rows_difference1 = []
    for index_name, row_name in tcdata['Test Case Name'].items():
        if not pd.isna(row_name):
            rows_with_nane.append(index_name+2)
        else:
            rows_with_nane1.append(index_name + 2)
    for index_no, row in tcdata[column_name].items():
        if not pd.isna(row):
            rows_with_step.append(index_no+2)
            rows_difference = set(rows_with_step).difference(set(rows_with_nane))
        else:
            rows_with_step1.append(index_no+2)
            rows_difference1 = set(rows_with_step1).difference(set(rows_with_nane1))
    if rows_difference:
        # print(column, '列:', rows_difference, '->去检查吧！')
        logger.error(f'首先检查TestCaseName是否为空!!!!如果CaseName存在的场合，请看左边->>>>> 项目 #{column_name}#{column}列:的第{rows_difference}->去检查吧.！')
    elif rows_difference1:
        logger.error(f'首先检查TestCaseName是否为空!!!!如果CaseName存在的场合，请看左边->>>>> #{column_name}#{column}列:的第{rows_difference1}->去检查吧！')
    time.sleep(1)


# 检查结果
def check_result(column, column_name):
    # print('开始')
    rows_nane = []
    illegality_row = []
    illegality_Set = []
    temdata = ['pass','passed','fail','failed','blocked']
    for index_name, row_name in tcdata['Test Case Name'].items():
        if not pd.isna(row_name):
            rows_nane.append(index_name+2)
    for illegality, row in tcdata[column_name].items():
        # print(row)
        rowLower = [s.lower() for s in [row] if isinstance(s, str)]
        # print(rowLower)
        if not set(rowLower).issubset(set(temdata)):
            illegality_row.append(illegality+2)
            # print(illegality_row)
    if illegality_row is not None:
        # print('列：', column, illegality_row, '字符不正确')
        logger.error(f'首先检查TestCaseName是否为空!!!!如果CaseName存在的场合，请看左边->>>>> 项目 #{column_name}#{column}列:的第{illegality_row}->字符不正确去检查吧！')
    time.sleep(1)


# 给单元格上色
def style_color(df):
    """

    :param df: pd.DataFrame
    :param colors: 字典 内容是 {标题:颜色}
    :return:
    """
    return df.style.apply(0, colors='#00EEEE')


# 0 check cover_name
def cv_name():
    if cv.name == '1.Cover_Changelog':
        logger.info(f'0.封面正确为:\t{cv.name}')
    elif cv.name != '1.Cover_Changelog':
        cv.name = '1.Cover_Changelog'
        logger.error(f'0.封面错误！！！！已修正为:\t{cv.name}')


# 1    Cover页面-阶段stage    按实际填写
def cv_stage():
    stage = cv.range('E2').value
    logger.info(f'1.测试阶段为:\t{stage}')


# 2    Cover页面-SW 版本    按实际填写，符合RTC defect填写要求。
def cv_sw_version():
    swv = cv.range('G2').value
    logger.info(f'2.测试版本为:\t{swv}')


# 3    Cover页面    Report 标题
def cv_title():
    title = cv.range('C3').value
    logger.info(f'3.报告标题为:\t{title}')


# 4    Cover页面    Report 参考文件信息
def cv_ref():
    for i in range(5,8):
        title = cv.range(f'B{i}:C{i}').value
        logger.info(f'4.报告引用为:\t{title}')


# 5    Cover页面    测试输入和环境信息
def cv_ref():
    for i in range(18,29):
        title = cv.range(f'B{i}:C{i}').value
        logger.info(f'5.报告输入为:\t{title}')


# 6    Issue List页面    NOK的项目必须列出，DefectID可以先不填
def add_sum(wb_name):
    # 去旧加新summary列表
    for sheet in wb.sheets:
        if 'Test Summary' in sheet.name:
            # print(sheet.name)
            sheet.delete()

    #wgy:增加一个参数wb_name
    md_name = wb_name.split('.')[0]
    # print(md_name)
    sumtitle = f'Test Summary{time.strftime("%Y%m%d_%H%M%S")}'
    wb.sheets.add(sumtitle, after='TestCase')
    sumsheet = wb.sheets[sumtitle]
    #模块名称放在第一列
    sumsheet.range('A1').value = 'Module Name'
    sumsheet.range('A2').value = md_name
    sumsheet.range('B1').value = 'Test Case Name'
    sumsheet.range('C1').value = 'Test Case Description'
    sumsheet.range('D1').value = 'DNG ID'
    sumsheet.range('E1').value = 'Result Overall State'
    sumsheet.range('F1').value = 'Result Details'


    sumsheet.range('B2').value = sumdata.values.tolist()
    rate_1 = sumdata.groupby(['Result Overall State']).size()
    # print(rate_1)
    rate_2 = pd.DataFrame(rate_1).reset_index()
    rate_2.set_axis(['State', 'Subtotal'], axis='columns', inplace=True)
    pass_rate = rate_2[rate_2['State'].isin(['blocked', 'passed', 'failed'])].reset_index(drop=True)
    # print(rate_2)
    rate_2['Percentage'] = rate_2['Subtotal']/rate_2['Subtotal'].sum()
    pass_rate['Percentage'] = pass_rate['Subtotal']/pass_rate['Subtotal'].sum()
    # print(rate_2)
    # print(pass_rate)
    sumsheet.range('H1').value = pass_rate
    sumsheet.range('I6').value = len(sumdata.index)
    sumsheet.autofit()
    logger.info(f'6.测试统计为:\t{sumtitle}')


# 7    Test Case Name    检查Test Case Name列，不允许出现重名case。使用excel条件格式查重。
def check_duplicate():
    global casecount
    casecount = tcdata['Test Case Name'].dropna().count()
    logger.info(f'7.用例总数为{casecount}!')

    # 检查重复名称
    casename = sumdata['Test Case Name']
    dupdata = casename[casename.duplicated()]
    # print(data1)
    if dupdata.values.tolist():
        logger.error(f'7.重复用例！！！！用例名为:\t{dupdata.values.tolist()}')
    else:
        logger.info(f'7.用例名称无重复!')

    # 检查名称长度
    for name in casename.values.tolist():
        # print(name)
        if len(name) != 14:
            logger.error(f'7.用例名称长度错误！！！！用例名为：\t{name}')

    # 检查名称间隔符
    for name in casename.values.tolist():
        # print(name)
        if name[6] != '_':
            logger.error(f'7.用例间隔符不为_！！！！用例名为：\t{name}')


# 8    Test Case Name    检查Test Case Name列，不允许出现case命名不符合要求的命名。
def rename_title():
    tc_col=tcdata.columns.values.tolist()
    tem_col=['Test Case Name', 'Model', 'Test Case Owner',
             'Test Case Priority', 'Test Case Description',
             'Test Case Functions', 'Test Case RequirementID',
             'Test Case UE-ID', 'Test Case Precondition',
             'Test Case Postcondition', 'Test Case Attachment', 'Step No.',
             'Step Action', 'Step Expected Result', 'Step Comment', 'Step Attachment',
             'Result Details', 'Result State', 'Result Overall State', 'Execution Start time',
             'Execution End time', 'Test Plan', 'Test PlanURI']
    if len(tc_col) != 23:
        logger.error(f'8.列数错误！相差:\t{23-len(tc_col)}')
    for i in range(len(tc_col)):
        if tc_col[i] != tem_col[i]:
            logger.error(f'8.第{get_char(i)}列名错误！！！！模板为{tem_col[i]}，实际为{tc_col[i]}')
        else:
            logger.info(f'8.第{get_char(i)}列名正确')


# 9    Model    检查Model列，不允许出现空的换行符，逗号等符号。
def check_model():
    global modelcount
    model_with_space = []
    model_with_line_character = []
    modelcount = tcdata['Model'].dropna().count()
    logger.info(f'9.Model总数为{casecount}!')
    if modelcount != casecount:
        logger.error(f'9.Model总数{modelcount}和Case总数{casecount}不相等!')
        check_count('B', 'Model')
    elif modelcount == casecount:
        # 判断存在空格
        modeldata = tcdata[['Model']].dropna()
        # print(modeldata)：
        a = modeldata[modeldata['Model'].str.contains(' ')]
        # print(a)
        b = modeldata[modeldata['Model'].str.endswith('\n')]
        if a.values.tolist():
            for index in a.index:
                name = tcdata.at[index, 'Test Case Name']
                logger.error(f'9.用例{name}的Model包含空格！！！！')
                model_with_space.append(index + 2)
        else:
            logger.info(f'9.Model正确：不含空格！')

        # 判断以换行符结尾
        if b.values.tolist():
            for index in b.index:
                name = tcdata.at[index, 'Test Case Name']
                logger.error(f'9.用例{name}的Model尾部有换行符！！！！')
                model_with_line_character.append(index + 2)
        else:
            logger.info(f'9.Model正确：尾部不含换行符！')
    if model_with_space:
        print('B列：', model_with_space, '有空格快去check')
    if model_with_line_character:
        print('B列：', model_with_line_character, '有换行符快去check')

# 10    Test Case Owner    均分配了正确的owner
def check_owner():
    global ownercount
    ownercount = tcdata['Test Case Owner'].dropna().count()
    logger.info(f'10.Test Case Owner总数为{ownercount}!')
    if ownercount != casecount:
        logger.error(f'10.Test Case Owner总数{ownercount}和Case总数{casecount}不相等!')
        check_count('C', 'Test Case Owner')

    #检查owner是否存在多个值
    ownersize = tcdata['Test Case Owner'].nunique()
    ownerdata = tcdata['Test Case Owner'].dropna().drop_duplicates().values.tolist()
    # print(ownersize)
    if ownersize > 1:
        logger.error(f'10.用例存在多个Owner，如下{ownerdata}')
    elif ownersize == 1:
        logger.info(f'10.Owner列正确')


# 11    Test Case Priority    均分配了正确的优先级.Basic的case优先级都是3
def check_prio():
    global priocount
    priocount = tcdata['Test Case Priority'].dropna().count()
    logger.info(f'11.Test Case Priority总数为{priocount}!')
    if priocount != casecount:
        logger.error(f'11.Test Case Priority总数{priocount}和Case总数{casecount}不相等!')
        check_count('D', 'Test Case Priority')

    # 检查prio是否存在多个值
    priosize = tcdata['Test Case Priority'].nunique()
    # print(priosize)
    priodata = tcdata['Test Case Priority'].dropna().drop_duplicates().values.tolist()
    # print(priodata)
    # priodata = map(int, priodata)
    if priosize > 1:
        logger.info(f'11.Test Case Priority存在多个优先级，如下{priodata}')
    elif priosize == 1:
        logger.info(f'11.Test Case Priority仅有1个优先级')

    if set(priodata).issubset([1, 2, 3]):
        logger.info(f'11.Test Case Priority字符正确')
    else:
        logger.error(f'11.Test Case Priority包含1/2/3之外的字符')


# 12    Test Case Description    均填写了正确的描述，一般复制需求文字即可。如是流程图的描述，可以自行编写。
def check_des():
    descount = tcdata['Test Case Description'].dropna().count()
    logger.info(f'12.Test Case Description数量为{descount}')
    if descount != casecount:
        logger.error(f'12.Test Case Description总数{descount}和Case总数{casecount}不相等!!!')
        check_count('E', 'Test Case Description')


# 13    Test Case Functions    填写function归属，仅填写二级功能即可。如HVAC_Vent，只需要填Vent
def check_fun():
    fuccount = tcdata['Test Case Functions'].dropna().count()
    logger.info(f'13.Test Case Functions数量为{fuccount}')
    if fuccount != casecount:
        logger.error(f'13.Test Case Functions总数{fuccount}和Case总数{casecount}不相等!!!')
        check_count('F', 'Test Case Functions')


# 14    Test Case Precondition    必须按照要求填写，包含关键标定、设备清单、设备状态、UI初始画面等。
def check_pre():
    precount = tcdata['Test Case Precondition'].dropna().count()
    logger.info(f'14.Test Case Precondition数量为{precount}')
    if precount != casecount:
        logger.error(f'14.Test Case Precondition总数{precount}和Case总数{casecount}不相等!!!')
        check_count('I', 'Test Case Precondition')


# 15    Test Case Postcondition    必须按照要求填写，包含关键标定、设备清单、设备状态、UI初始画面等。
def check_post():
    postcount = tcdata['Test Case Postcondition'].dropna().count()
    logger.info(f'15.Test Case Postcondition数量为{postcount}')
    if postcount != casecount:
        logger.error(f'15.Test Case Postcondition总数{postcount}和Case总数{casecount}不相等!')
        check_count('J', 'Test Case Postcondition')


# 16    Step No.    必填，并且必须有值、可以是1,2,3  或者step 1，step 2，step 3
def check_step_no():
    # stepcount = tcdata['Step No.'].dropna().count()
    # print(stepcount)
    nullcount = tcdata['Step No.'].isnull().sum()
    # print(nullcount)
    logger.info(f'16.Step No.空值数量为{nullcount}')
    if casecount != nullcount+1:
        logger.error(f'16.Step No.包含非法空值！！！！')
        check_none('L', 'Step No.')



# 17    Step Action    必填
def check_step_action():
    nullcount = tcdata['Step Action'].isnull().sum()
    # print(nullcount)
    logger.info(f'17.Step Action空值数量为{nullcount}')
    if casecount != nullcount+1:
        logger.error(f'17.Step Action包含非法空值！！！！')
        check_none('F', 'Step Action')


# 18    Step Expected Result    必填
def check_step_result():
    nullcount = tcdata['Step Expected Result'].isnull().sum()
    # print(nullcount)
    logger.info(f'18.Step Expected Result空值数量为{nullcount}')
    if casecount != nullcount+1:
        logger.error(f'18.Step Expected Result包含非法空值！！！！')
        check_none('M', 'Step Action')


# 19    Result Details    非OK项目，必须填写detail信息，NG项目保留视频、截图信息
# def check_step_detail():
#     postcount = tcdata['Test Case Postcondition'].dropna().count()
#     logger.info(f'11.功能类别正确，数量为{postcount}')


# 20    Result State    Step result必填，使用：pass/passed   fail/failed  blocked   可以全大写
def check_step_state():
    nullcount = tcdata['Result State'].isnull().sum()
    # print(nullcount)
    logger.info(f'20.Result State空值数量为{nullcount}')
    if casecount != nullcount+1:
        logger.error(f'20.Result State包含非法空值！！！！')
        check_none('R', 'Result State')
    statedata = tcdata['Result State'].dropna().drop_duplicates().values.tolist()
    statedatalower = [x.lower() for x in statedata]
    temdata = ['pass','passed','fail','failed','blocked','deferred']
    if set(statedatalower).issubset(set(temdata)):
        logger.info(f'20.Result State字符正确')
    else:
        logger.error(f'20.Result State包含{temdata}之外的字符(不区分大小写)')
        check_result('R', 'Result State')

# 21    Result Overall State    Case result必填，使用：pass/passed   fail/failed  blocked ，只有全部step pass，才认为case pass
def check_step_overall():
    postcount = tcdata['Result Overall State'].dropna().count()
    logger.info(f'21.Result Overall State数量为{postcount}')
    if postcount != casecount:
        logger.error(f'21.Result Overall State总数{postcount}和Case总数{casecount}不相等!')
        check_count('S', 'Result Overall State')
    overalldata = tcdata['Result Overall State'].dropna().drop_duplicates().values.tolist()
    soveralldatalower = [x.lower() for x in overalldata]
    temdata = ['pass','passed','fail', 'failed', 'blocked', 'deferred']
    if set(soveralldatalower).issubset(set(temdata)):
        logger.info(f'21.Result Overall State字符正确')
    else:
        logger.error(f'21.Result Overall State包含{temdata}之外的字符(不区分大小写)')
        check_result('S', 'Result Overall State')


# 22    Test Plan    填写RQM上已创建的Plan名称，只需要在第2行，填一次。
def check_plan():
    title = tc.range('V2').value
    logger.critical(f'22.Test Plan为:\t[{title}]')
    titlelist = ['IVB','IVER', '60BOR', '80BOR','PPV','VTC','NS','S','20.5PPV','20.5NS','20.5S']
    # if title not in titlelist:
    #     #     logger.error(f'22.Test Plan不属于{titlelist}')


# 23    Test PlanURI    "填写RQM上已创建的Plan ID URI，只需要在第2行，填一次。
# 模板：urn:com.ibm.rqm:testplan:658其中658是对应的Plan在RQM上的ID"
def check_planlink():
    title = tc.range('W2').value
    logger.critical(f'23.Test PlanURI为:\t[{title}]')
    if title[0:25] != 'urn:com.ibm.rqm:testplan:':
        # print(title[0:24])
        logger.error(f'23.Test PlanURI前缀错误，不为urn:com.ibm.rqm:testplan:')

if __name__ == '__main__':
    print(f'当前目录下文件清单为{get_xlsx()}')
    init()
    for x in get_xlsx():
        write_log(x)
        read_data(x)
        cv_name()
        cv_stage()
        cv_sw_version()
        cv_title()
        cv_ref()
        add_sum(x)
        check_duplicate()
        rename_title()
        check_model()
        check_owner()
        check_prio()
        check_des()
        check_fun()
        check_pre()
        check_post()
        check_step_no()
        check_step_action()
        check_step_result()
        check_step_state()
        check_step_overall()
        check_plan()
        check_planlink()
        save_quit()
    time.sleep(5)
    app.quit()
