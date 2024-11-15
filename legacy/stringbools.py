import re

def hasNumbers(s):
    return any(i.isdigit() for i in s)

def hasLetters(s):
    return re.search('[a-zA-Z]', s)

def hasYear(s):
    numbercount = 0
    for i in s:
        if i.isdigit():
            numbercount = numbercount + 1
    if numbercount > 3:
        return True
    else:
        return False

def isAllLower(s):
    nonlowercase = 0
    for i in s.replace(' ', ''):
        if not i.islower():
            nonlowercase = nonlowercase + 1
            break
    if nonlowercase > 0:
        return False
    else:
        return True