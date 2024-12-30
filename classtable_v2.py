import pandas as pd
import random
import math
import time
from matplotlib import pyplot as plt
import sys

# 读取Excel中的表格数据
courses = pd.read_excel(r'./data.xlsx', sheet_name="courses")
rooms = pd.read_excel(r'./data.xlsx', sheet_name="rooms")
teachers = pd.read_excel(r'./data.xlsx', sheet_name="teachers")
classes = pd.read_excel(r'./data.xlsx', sheet_name="classes")
schedule = pd.read_excel(r'./data.xlsx', sheet_name="schedule")
classmap = pd.read_excel(r'./data.xlsx', sheet_name="classmap")
schedule.set_index('教室', inplace=True)
rooms.set_index('教室', inplace=True, drop=False)
classes.set_index('班级', inplace=True, drop=False)

# 创建任务集合
def task_set():
    task = pd.DataFrame([], columns=['课程编号', '班级', '老师', '人数'])
    for i in courses.index:
        for j in teachers.columns:
            if courses['课程编号'][i] in list(teachers[j]):
                t = list()
                t.append(courses['课程编号'][i])
                t.append(courses['班级'][i])
                t.append(j)
                t.append(classes['人数'][courses['班级'][i]])
                task.loc[len(task.index)] = t
    return task

# 计算当前班级有哪些课，即属于该班级的task条目索引
def task_course_class(task, c):
    course = set()
    for i in task.index:
        if task['班级'][i] == c:
            course.add(i)
    return course

# 计算课表是否有冲突，是否有同一个班级同时上两节课
def whether_conflict_class(task):
    for i in schedule.columns:
        for j in classes['班级']:
            if len(set(schedule[i]).intersection(task_course_class(task, j))) > 1:
                return True
    return False

# 计算当前老师教哪些课，即属于该老师的task条目索引
def task_course_teacher(task, t):
    course = set()
    for i in task.index:
        if task['老师'][i] == t:
            course.add(i)
    return course

# 计算课表是否有冲突，是否有同一位老师同时上两节课
def whether_conflict_teacher(task):
    for i in schedule.columns:
        for j in teachers.columns:
            if len(set(schedule[i]).intersection(task_course_teacher(task, j))) > 1:
                return True
    return False

# 贪心算法生成初始总课表
def init_schedule(task):
    task.sort_values(by='人数', inplace=True)
    task.reset_index(inplace=True, drop=True)
    rooms.sort_values(by='容量',inplace=True)
    rooms.set_index('教室', inplace=True, drop=False)
    for i in task.index:
        for j in rooms.index:
            if (rooms['容量'][j] >= task['人数'][i]) and ('空' in list(schedule.loc[j])):
                for k in schedule.columns:
                    if schedule[k][j] == '空':
                        schedule[k][j] = i
                        if whether_conflict_class(task) or whether_conflict_teacher(task):
                            schedule[k][j] = '空'
                        else:
                            break
                if i in list(schedule.loc[j]):
                    break

# 利用软约束对课表评分
def evaluate_schedule(task, temp_schedule):
    score = 0
    for i in temp_schedule.columns:
        for j in temp_schedule.index:
            if temp_schedule[i][j] != '空':
                score += (task['人数'][temp_schedule[i][j]] / rooms['容量'][j])
                if i[-1] in ['2', '4']:
                    score += 1
                if i[-1] == '5':
                    score -= 0.5
    return score

# 模拟退火算法
def simulated_annealing(task, temperature=100, cooling_rate=0.95, min_temperature=1, inner_loop=100):
    path = list()
    current_schedule = schedule.copy(deep=True)
    current_score = evaluate_schedule(task, current_schedule)
    best_schedule = current_schedule.copy(deep=True)
    best_score = current_score
    while temperature > min_temperature:
        for i in range(inner_loop):
            # 随机选择一个教室和一个时间的课程
            time1 = random.choice(schedule.columns)
            room1 = random.choice(schedule.index)
            # 随机选择另一个教室和另一个时间的课程
            time2 = random.choice(schedule.columns)
            room2 = random.choice(schedule.index)
            # 若两次随机选择同一节课则重新选择
            if (room1 == room2) and (time1 == time2):
                continue
            # 检测本次交换是否合理是否有冲突
            if whether_conflict_class(task) or whether_conflict_teacher(task):
                continue
            # 交换课程
            new_schedule = current_schedule.copy(deep=True)
            new_schedule[time1][room1] = current_schedule[time2][room2]
            new_schedule[time2][room2] = current_schedule[time1][room1]
            new_score = evaluate_schedule(task, new_schedule)
            # 判断是否接受新的课程表
            delta = current_score - new_score
            if delta < 0 or random.random() < math.exp(-delta/temperature):
                current_schedule = new_schedule.copy(deep=True)
                current_score = new_score
                # 记录搜索路径
                path.append(current_score/(84*2))
                if current_score > best_score:
                    best_schedule = current_schedule.copy(deep=True)
                    best_score = current_score
        # 降温
        temperature *= cooling_rate
    return best_schedule, path

# 总课表
def create_total_schedule(task, best_schedule):
    total_schedule = best_schedule.copy(deep=True)
    total_courses = courses.copy(deep=True)
    total_courses.set_index('课程编号', inplace=True, drop=False)
    for i in total_schedule.columns:
        for j in total_schedule.index:
            k = total_schedule[i][j]
            if k != '空':
                total_schedule[i][j] = total_courses['课程'][task['课程编号'][k]] + '|' + task['班级'][k] + '|' + task['老师'][k]
    return total_schedule

# 单个班级
def create_single_schedule(task, best_schedule, total_schedule):
    n = 0
    for c in classes['班级']:
        single_schedule = best_schedule.copy(deep=True)
        for i in single_schedule.columns:
            for j in single_schedule.index:
                k = single_schedule[i][j]
                if k != '空':
                    if task['班级'][k] == c:
                        single_schedule[i][j] = total_schedule[i][j]
                    else:
                        single_schedule[i][j] = '空'
        final_schedule = pd.DataFrame([], columns=['Mon', 'Tue', 'Wed', 'Thu', 'Fri'], index=['第1节', '第2节', '第3节', '第4节', '第5节'])
        day = 0
        section = 0
        for t in single_schedule.columns:
            for h in single_schedule.index:
                if single_schedule[t][h] != '空':
                    final_schedule.iloc[day, section % 5] = single_schedule[t][h] + '|' + h
                    break
            else:
                final_schedule.iloc[day, section % 5] = '空'
            section += 1
            if section % 5 == 0:
                day += 1
        for tmp in classmap[c]:
            if type(tmp) != float:
                mode = 'w'
                if n:
                    mode = 'a'
                with pd.ExcelWriter('single_schedule.xlsx', engine='openpyxl', mode=mode) as writer:
                    final_schedule.to_excel(writer, sheet_name=tmp)
                n += 1 

# 折线图
def curve(path):
    plt.plot(range(1,len(path)+1), path, alpha=0.5, linewidth=1, label='acc')
    plt.legend()  #显示上面的label
    plt.xlabel('次数') #x_label
    plt.ylabel('number')#y_label
    plt.show()

# 主函数
if __name__ == '__main__':
    # 计算代码运行时间
    time_start = time.perf_counter()

    task = task_set()
    init_schedule(task)
    tanxin_score = evaluate_schedule(task, schedule)
    print(tanxin_score)
    # 测试
    # tmp = int(sys.argv[1])
    
    best_schedule, path = simulated_annealing(task, temperature=1000, cooling_rate=0.96, min_temperature=1, inner_loop=100)
    # print(path)
    curve(path)
    best_schedule.to_excel('best_schedule.xlsx', sheet_name='Sheet1')
    best_score = evaluate_schedule(task, best_schedule)
    total_schedule = create_total_schedule(task, best_schedule)
    print(best_score)
    print(((best_score - tanxin_score)/tanxin_score) * 100, '%')
    total_schedule.to_excel('total_schedule.xlsx', sheet_name='Sheet1')
    create_single_schedule(task, best_schedule, total_schedule)

    time_end = time.perf_counter()
    time_sum = time_end - time_start
    print('time:', time_sum, 's')

