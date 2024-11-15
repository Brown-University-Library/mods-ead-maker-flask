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