
# 使用须知：
# 默认中国人和韩国人人数一样了哦
# 一共有2个路径要改
# 可以自己改参数，在def matchPoint()那里

import pandas as pd


class person():
    def __init__(self, name, email, gender, pGender, language, hobby, sentence):
            self.name = name
            self.email = email
            self.gender = gender
            self.pGender = pGender
            self.language = language
            self.hobby = hobby
            self.sentence = sentence
    


# 读取Excel文件，1是中国，2是韩国，自己改下路径
excel_file_path1 = r'C:\xxx\xxx\match1.xlsx'
df1 = pd.read_excel(excel_file_path1)
excel_file_path2 = r'C:\xxx\xxx\match2.xlsx'
df2 = pd.read_excel(excel_file_path2)


num_rows = df1.shape[0]
num_colonm = df1.shape[1]
chi = []
kor = []
chiAnswer = [1, 3, 3, 2, 2]
korAnswer = [3, 1, 2, 2, 1] 

for i in range(num_rows):
      row = df1.iloc[i]
      lan = 0
      for j in range(8, 13):
            if row[j] == chiAnswer[j - 8]:
                  lan += 1
      ho = row[19: num_colonm - 1]
      chi.append(person(row[6], row[13], row[14], row[15: 18], lan, ho, row[num_colonm-1]))

for i in range(num_rows):
      row = df2.iloc[i]
      lan = 0
      for j in range(8, 13):
            if row[j] == korAnswer[j - 8]:
                  lan += 1
      ho = row[19: num_colonm-1]
      kor.append(person(row[6], row[13], row[14], row[15: 18], lan, ho, row[num_colonm-1]))
      

# a挑b
def matchPoint(a, b):
      point = 0
      if a.pGender[b.gender-1] == 1:
            point += 600
      point += abs(a.language + b.language - 6) * 10
      for i in range(37):
            if a.hobby[i] == 1 and b.hobby[i] == 1:
                  point += 10
            if a.hobby[i] == 0 and b.hobby[i] == 0:
                  point += 2
      return point

def point(a):
      return a[1]


def order(a, bList):
      pointList = []
      for i in bList:
            pointList.append([i, matchPoint(a, i)])
      sorted_indices = sorted(range(len(pointList)), key=lambda i: pointList[i][1], reverse=True)
      
      return sorted_indices

      

      
      









# below is copied from https://www.geeksforgeeks.org/stable-marriage-problem/

# Python3 program for stable marriage problem

# Number of Men or Women
N = num_rows

# This function returns true if
# woman 'w' prefers man 'm1' over man 'm'
def wPrefersM1OverM(prefer, w, m, m1):
	
	# Check if w prefers m over her
	# current engagement m1
	for i in range(N):
		
		# If m1 comes before m in list of w,
		# then w prefers her current engagement,
		# don't do anything
		if (prefer[w][i] == m1):
			return True

		# If m comes before m1 in w's list,
		# then free her current engagement
		# and engage her with m
		if (prefer[w][i] == m):
			return False

# Prints stable matching for N boys and N girls.
# Boys are numbered as 0 to N-1.
# Girls are numbered as N to 2N-1.
def stableMarriage(prefer):
	
	# Stores partner of women. This is our output
	# array that stores passing information.
	# The value of wPartner[i] indicates the partner
	# assigned to woman N+i. Note that the woman numbers
	# between N and 2*N-1. The value -1 indicates
	# that (N+i)'th woman is free
	wPartner = [-1 for i in range(N)]

	# An array to store availability of men.
	# If mFree[i] is false, then man 'i' is free,
	# otherwise engaged.
	mFree = [False for i in range(N)]

	freeCount = N

	# While there are free men
	while (freeCount > 0):
		
		# Pick the first free man (we could pick any)
		m = 0
		while (m < N):
			if (mFree[m] == False):
				break
			m += 1

		# One by one go to all women according to
		# m's preferences. Here m is the picked free man
		i = 0
		while i < N and mFree[m] == False:
			w = prefer[m][i]

			# The woman of preference is free,
			# w and m become partners (Note that
			# the partnership maybe changed later).
			# So we can say they are engaged not married
			if (wPartner[w - N] == -1):
				wPartner[w - N] = m
				mFree[m] = True
				freeCount -= 1

			else:
				
				# If w is not free
				# Find current engagement of w
				m1 = wPartner[w - N]

				# If w prefers m over her current engagement m1,
				# then break the engagement between w and m1 and
				# engage m with w.
				if (wPrefersM1OverM(prefer, w, m, m1) == False):
					wPartner[w - N] = m
					mFree[m] = True
					mFree[m1] = False
			i += 1

			# End of Else
		# End of the for loop that goes
		# to all women in m's list
	# End of main while loop
	return wPartner
	'''
	# Print solution
	print("Woman ", " Man")
	for i in range(N):
		print(i + N, "\t", wPartner[i])
    '''
    
                
	


# Driver Code







prefer = []
for chiP in chi:
       prefer.append(order(chiP, kor))
for korP in kor:
       prefer.append(order(korP, chi))

matchList = stableMarriage(prefer)

#print




#这是参与名单直接输出
print("分组名单")
for i in range(num_rows):
       print(i+1, kor[i].name, " & ", chi[matchList[i]].name)

#这是用来发邮件的直接输出
print("发邮件")
for i in range(num_rows):
       print("第", i+1, "组")
       print(kor[i].email, chi[matchList[i]].email)
       print("以下是你&你的语伴的信息哦：")
       print(kor[i].name, " & ", chi[matchList[i]].name)
       print("你们送给互相的一句话是：")
       print(kor[i].sentence)
       print(chi[matchList[i]].sentence)
       
       print("아래는 여러분과 여러분의 어학 파트너의 정보입니다!")
       print(kor[i].name, " & ", chi[matchList[i]].name)
       print("서로에게 한마디!")
       print(kor[i].sentence)
       print(chi[matchList[i]].sentence)
       

