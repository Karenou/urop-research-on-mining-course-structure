import xlwings as xw
from math import sqrt
import operator
import time, random
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import skfuzzy as fuzzy


#fuzzy c means parameters
# epsilon = 0.00000001
# max = 10000
# m = 1.5   # 1.5-2.5


def grade(x):
    return {
        'P': 1.0,
        'PP': 1.0,          #try to remove the courses with P grade
        'F':1.0,
        'D': 1.3,
        'C-': 1.7,
        'C': 2.0,
        'C+': 2.3,
        'B-': 2.7,
        'B': 3.0,
        'B+': 3.3,
        'A-': 3.7,
        'A': 4.0,
        'A+': 4.3
    }.get(x, 0.0)


def spearman_corr(neighbor, target):
    similarity = {}
    for user_index, user_rating in enumerate(neighbor):
        sum_squared = 0.0
        rated_item= 0
        for course_index, course_grade in enumerate(user_rating):
            if user_rating[course_index] != 0 and target[course_index] != 0:
                rated_item += 1
                sum_squared += pow(user_rating[course_index] - target[course_index], 2)
            else:
                continue
        if rated_item != 0:
            similarity[user_index] = 1 - 6 * sum_squared / (rated_item * (pow(rated_item, 2) - 1))
        else:
            similarity[user_index] = 0.0

    similarity = sorted(similarity.items(), key=operator.itemgetter(1), reverse=True)
    # print("Find 100 similar students with time: " + str(round(time.time() - start, 2)) + " seconds")
    return similarity[:100]


def cos_similarity(neighbor,target):
    # find the common courses other and target students both take
    similarity = {}

    for user_index, user_rating in enumerate(neighbor):
        numerator = 0.0
        sum_sq_target = 0.0
        sum_sq_other = 0.0
        for course_index, course_grade in enumerate(user_rating):
            if user_rating[course_index] != 0 and target[course_index] != 0 :
                numerator += user_rating[course_index] * target[course_index]
                sum_sq_target += pow(target[course_index],2)
                sum_sq_other += pow(user_rating[course_index],2)
            else:
                continue

            if sqrt(sum_sq_target * sum_sq_other) != 0:
                similarity[user_index] = numerator / sqrt(sum_sq_target * sum_sq_other)
            else:
                similarity[user_index] = 0.0

    similarity = sorted(similarity.items(),key=operator.itemgetter(1),reverse=True)
    # print("Find 100 similar students with time: " + str(round(time.time() - start, 2)) + " seconds")
    return similarity[:100]


def mae(predict, actual):
    if (actual - predict) >= 0:
        return actual - predict
    else:
        return predict - actual


def read_data(admit_type1):    # change according to the number of admission type
    wb = xw.Book("StudentCourseEnrollmentData2.xlsx")
    a = wb.sheets[0]
    b = wb.sheets[1]

    PassFail= ['P','PP','DI','**','PA','W','AU','I']
    course_data = a['A1:K394145'].options(pd.DataFrame, index=False, header=True).value
    student_info = b['A1:Y11362'].options(pd.DataFrame, index=False, header=True).value

    df_student_info = student_info.loc[(student_info['Admit Type Desc'] == admit_type1)]
    stdList = df_student_info['Masked ID'].astype(int)
    admission = df_student_info['Admit Type Desc']
    average_grade = df_student_info['CGA']
    student_profile = pd.DataFrame({'students': stdList, 'CGA':average_grade, 'admission type': admission})
    stdList = stdList.tolist()
    student_profile.to_csv('data\student_CGA_local_direct_entry.csv',header=True)       # change according to the file name

    df_course_data1 = course_data.loc[course_data['Masked ID'].isin(stdList)]
    df_course_data = df_course_data1.loc[df_course_data1['Grade'].isin(PassFail) == False]    # bug1

    stdid = df_course_data['Masked ID'].astype(int)
    df_courses = df_course_data['Subject'] + df_course_data['Catalog']
    df_grade = df_course_data['Grade'].apply(grade)
    course_history = pd.DataFrame({'Grade': df_grade, 'courses': df_courses, 'students': stdid})
    course_history.to_csv('data\course_history_local_direct_entry.csv', header=True)              # change according to the file name


def initialize(start, num_training, admit_type):
    # import data
    course_history_path = 'data\course_history_' + admit_type + '.csv'
    student_CGA_path = 'data\student_CGA_' + admit_type + '.csv'
    course_history = pd.read_csv(course_history_path,dtype={'Grade':float, 'students':str, 'courses':str})
    student_CGA= pd.read_csv(student_CGA_path, dtype={'students': str, 'CGA': float})

    # store all students' data {SID1: { course1: grade1, course2: grade2 ... }, SID2: {} ... }
    # exclude courses that are graded based on pass or fail
    data = {}
    stdList = course_history['students'].tolist()
    course_list = course_history['courses'].tolist()
    for i in course_history.index:
        temp_SID = stdList[i]
        if temp_SID not in data:
            data[temp_SID] = {}
        temp_course = course_list[i]
        data[temp_SID][temp_course] = course_history['Grade'][i]

    # count number of courses and students
    df_courses = pd.read_csv('data\course_history.csv',dtype= {'courses':str})
    courses = df_courses['courses'].drop_duplicates()
    courses = courses.tolist()
    users = course_history['students'].drop_duplicates()
    users = users.tolist()

    pip_prep = pd.DataFrame(columns=['courses', 'min', 'max', 'medium','sum','num_rated','average'])
    pip_prep['course']= courses
    pip_prep.fillna(value={'min': 0.0, 'max': 0.0, 'medium': 0.0, 'num_rated': 0, 'average':0.0})

    # delete the student records of less than 5 courses
    for v in users:
        x = len(data[v].keys())
        if x < 5 :
            del data[v]
            del users[users.index(v)]

    # store as rating matrix
    all_data = []
    for i in users:
        temp = [0.0] * len(courses)
        for j in range(0,len(courses)):
            try:
                if courses[j] in data[i].keys():
                    temp[j] = float(data[i][courses[j]])
                    # fill the pip_pre
                    index = pip_prep['course']==courses[j]
                else:
                    temp[j] = 0.0
            except:
                print(j, courses[j])
                print(i, data[i])
        all_data.append(temp)

    target_data = [[0.0] * len(courses)] * num_training # split the data into train and testing (first 10 students)
    target_CGA = student_CGA['CGA'][:num_training]
    target_SID = users[:num_training]

    predict_course = [(lambda v: data[v].keys()[random.randint(0,10)]if len(data[v].keys()) > 10 else data[v].keys()[random.randint(0,5)])(x) for x in target_SID]

    # store the predicted student course history in {SID: {course1: grade, course2: grade }}
    for j in range(0, len(target_SID)):
        for i in course_history.index:
            if stdList[i] == target_SID[j]:
                course = course_history['courses'][i]
                course_grade = course_history['Grade'][i]
                target_data[j][courses.index(course)] = course_grade

    print("Store all student data using " + str(round(time.time() - start, 2)) + " seconds.")
    print

    return all_data, target_data, target_SID, target_CGA, users, courses, predict_course

# find intersection / union
def Jaccard(all_data, target, hurdle):
    jaccard = [0.0] * len(all_data)

    for i in range(0, len(all_data)):
        intersection = 0
        setA = 0
        setB = 0
        for j in range(0, len(target)):
            if all_data[i][j] != 0.0:
                setA += 1
                if target[j] != 0.0:
                    intersection += 1
            if target[j] != 0.0:
                setB += 1
        if (setA + setB - intersection) == intersection:   #neglect target student himself
            jaccard[i] = 0.0
        elif (setA + setB - intersection) > 0:
            jaccard[i] = intersection / float(setA + setB - intersection)
        else:
            jaccard[i] = 0.0

    # jaccard.sort(reverse=True)
    # print(jaccard)
    # select users
    neighbor = []
    for i in range(0, len(jaccard)):
        if jaccard[i] > hurdle:
            neighbor.append(all_data[i])

    return neighbor


def fuzzycmeans(data, num_cluster,m,error, maxiter, courses, users):
    # cluster center and fuzzy c partitioned matrix
    center = [[0.0 for i in range(0, len(courses))] for i in range(0, num_cluster)]
    u = [[0.0 for i in range(0, len(courses))] for i in range(0, len(users))]
    center, u,u0, d, jm, p, fpc = fuzzy.cluster.cmeans(data,num_cluster,m,error,maxiter, init=None, seed=0)

    return center, u



# main code

hurdle = 0.26
num_training = 200
RMSE  = 0.0
mean_sq_error = [0.0] * num_training

# 1. International, Mainland/Taiwan/Macau (A-level, IB); 2. JUPAS; 3. JEE (Gaokao); 4. Local direct entry
admit_type = ['UG International', 'UG Mainland/Taiwan/Macau', 'UG JUPAS of 4-year cohort','UG JUPAS','UG JEE', 'UG Local Direct Entry']
hurdle_rates = {'all': 0.26, 'JEE': 0.1, 'International':0.26, 'JUPAS': 0.26, 'local_direct_entry' :0.06}
# read_data(admit_type[2], admit_type[3])
# read_data(admit_type[5])

# check how much time consumed
start = time.time()

# data initialization
all_data,target_data,target_SID, target_average_grade, users, courses, predict_course = initialize(start, num_training, "JUPAS")

for j in range(0, len(target_data)):
    numerator = 0.0
    sum_sim = 0.0
    # filter the neigbour capped at 0.26 Jaccard coefficient
    neighbor = Jaccard(all_data, target_data[j], hurdle)
    similarity = cos_similarity(neighbor, target_data[j])  # top 100 similar users

    for i in range(0, len(similarity)):
        user_index = similarity[i][0]
        course_index = courses.index(predict_course[j])
        if neighbor[user_index][course_index] != 0.0:
            numerator += (neighbor[user_index][course_index] - target_average_grade[j]) * similarity[i][1]
            sum_sim += abs(similarity[i][1])
    if sum_sim != 0.0:
        score = target_average_grade[j] + numerator / sum_sim
    else:
        score = target_average_grade[j]

    # neighbor = np.array(neighbor)
    # center, u = fuzzycmeans(neighbor,5,m,epsilon, max,courses, users)

    # print "Student " + target_SID[j] + ", your predicted grade for course " + predict_course[j] + " is: " + str(round(score, 4))
    # if score <= 1.3:
    #     print "D or Fail"
    # elif score <= 2.5:
    #     print "C Range"
    # elif score <= 3.5:
    #     print "B Range"
    # else:
    #     print "A Range"
    #
    # print ("Your actual grade is:" + str(target_data[j][courses.index(predict_course[j])]))           # BUG
    mean_sq_error[j] = mae(score, target_data[j][courses.index(predict_course[j])])
    RMSE += pow(mean_sq_error[j],2)


print("Total time used is: " + str(round(time.time() - start,2)) + " seconds")
print("RMSE is: ", sqrt(RMSE / num_training))

# x = np.arange(0,num_training,1)
# error = np.array(mean_sq_error)
# plt.ylabel("MAE")
# plt.plot(x, error, 'ro')
# plt.axis([0,num_training,0,5])
# plt.show()


