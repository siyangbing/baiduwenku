import itertools
import time

word_str = "无裂纹,无松动|无突起,无异物|无划痕|无锈蚀,是否完好|是否到位|是否牢固|是否正常"

def word_split(word_str):
    #按','分割字符
    all_word_list = word_str.split(',')
    str_result = ""
    # 对分割好的每一个字符做排列
    for one_word_list in all_word_list:
        # 按'|'分割字符
        one_word = one_word_list.split("|")
        # 取全排列
        for i in range(1, len(one_word)):
            #取排列组合
            for word in itertools.permutations(one_word, i):
                word_str = ''
                for w1 in word:
                    word_str += w1 + ','
                # print(word_str)
                str_result += word_str[:-1] + '|'
    return str_result

t0 = time.time()
rr_result = word_split(word_str)
print(rr_result)
print(time.time()-t0)
