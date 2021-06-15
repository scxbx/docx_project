#!/usr/bin/python
# -*- coding: UTF-8 -*-

import re


# 四个分隔符为：,  ;  *  \n

def split_name(organization, xiang_or_zhen):
    level1 = '县'
    level2 = xiang_or_zhen

    re_str = '{}|{}'.format(level1, level2)
    return re.split(re_str, organization)


def xiang_or_zhen(organization):
    if '乡' in organization:
        return '乡'
    elif '镇' in organization:
        return '镇'
    else:
        print("既不是乡也不是镇")
        return ''


if __name__ == '__main__':

    b = '澄迈县金江镇博潭村委会长坡仔村民小组'
    result = split_name(b, xiang_or_zhen(b))
    print(result)

    print('{}县{}{}{}'.format(result[0], result[1], xiang_or_zhen(b), result[2]))


