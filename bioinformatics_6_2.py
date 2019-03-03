# smith-waterman algorithm 

f = raw_input('Please enther the first nucleotide sequence: ')
s = raw_input('Please enter the second nucleotide sequence: ') 

matrix = []
len_f = len(f)+1 
len_s = len(s)+1 

matrix = [[0 for j in range(len_f)] for i in range(len_s)]
matrix2 = [['#' for j in range(len_f)] for i in range(len_s)] 

# mismatch _score 
d= -1
# match score 
m = 3
# gap penalty 
g = -2 

match_s=0

for i in range(1,len_s): 
    for j in range(1,len_f): 
        try: 
            if f[j-1:j] == s[i-1:i]: 
                match_s = m
            else: match_s = d
        except: 
            continue

        match = matrix[i-1][j-1] + match_s 
        match2 = matrix[i-1][j] + g 
        match3 = matrix[i][j-1] + g 

        matrix[i][j] = max(match, match2, match3, 0)
        if matrix[i][j] == match: matrix2[i][j] = 'm' 
        elif matrix[i][j] == match2: matrix2[i][j] = 'm2'
        elif matrix[i][j] == match3: matrix2[i][j] = 'm3'


align_a = ''
align_b = '' 

max_num = 0
index_f = 0 
index_s = 0 
for i in range(1, len_s): 
    for j in range(1, len_f): 
        if matrix[i][j] > max_num: 
            max_num = matrix[i][j]
            index_f = j 
            index_s = i 

lf = index_f
ls = index_s

while lf >= 0 and ls >= 0: 
    if lf >= 0 and ls >= 0 and matrix2[ls][lf] == 'm': 
        align_a += f[lf-1:lf]
        align_b += s[ls-1:ls]
        lf -=1 
        ls -=1
    elif ls >= 0 and matrix2[ls][lf] == 'm3': 
        align_a += f[lf-1:lf]
        align_b += '-' 
        lf -= 1 
    else:  
        align_a += '-' 
        align_b += s[ls-1:ls]
        ls -= 1 


print 'Score = %s' % max_num 
print align_a[-2::-1] 
print align_b[::-1]
